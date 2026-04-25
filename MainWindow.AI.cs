namespace AbacaWpf;

public partial class MainWindow
{
    // ИИ вызывает эту проверку после остановки каждого броска.
    // Метод решает, стоит ли уже записывать найденный ход или лучше продолжать бросать:
    // готовые сильные комбинации останавливаются рано, а спорные суммы/каре могут ждать улучшения.
    private bool ShouldComputerStop(ComputerMove best)
    {
        var threshold = CurrentPlayer.Difficulty switch
        {
            ComputerDifficulty.Careful => 20,
            ComputerDifficulty.Aggressive => 42,
            _ => 32
        };

        if (best.Score <= 0)
            return false;

        if (best.Row == 14 && _rollCount == 1 && IsGoodOpeningSum())
            return true;

        if (best.Row == 14 && _rollCount < 3)
            return false;

        if (best.Row == 12 && _rollCount == 2 && CanChaseAbaca())
            return false;
        if (best.Row == 12 && !IsKareWorthRecording(GetDiceCounts()))
            return false;
        if (best.Row == 9 && !IsFullHouseWorthRecording(GetDiceCounts()))
            return false;

        if (_rollCount == 1 && best.Score > 0 && IsTacticalPrizeMove(best.Row))
            return true;

        if (_rollCount == 1)
            return best.Row == 13
                || best.Row == 12 && best.Score > 0
                || best.Row == 6 && best.Score >= 8 && (IsTacticalPrizeMove(6) || GetTacticalPrizeUrgency(6) >= 140)
                || best.Row == 7 && (HasPremiumTwoPairs(GetDiceCounts()) || IsTacticalPrizeMove(7) || ShouldStopOnOpeningTwoPairs(GetDiceCounts(), best.Score))
                || best.Row == 8 && best.Score > 0 && (IsTacticalPrizeMove(8) || IsStrongOpeningTriple(GetDiceCounts()) && !HasStraightFork())
                || best.Row == 9 && best.Score > 0
                || best.Row is 10 or 11 && best.Score > 0;

        if (best.Row is 10 or 11 && best.Score > 0)
            return true;

        if (_rollCount >= 2 && best.Score > 0 && ShouldStopOnCompletedTacticalPrizeMove(best.Row))
            return true;

        if (best.Row == 7 && best.Score > 0 && _rollCount >= 2 && IsTacticalPrizeMove(7))
            return true;

        if (best.Row == 9
            && _rollCount >= 2
            && (best.Score >= 24
                || CountBusyMainCellsInRow(9) <= 2
                || (ColumnCount - 1 - CountBusyMainCellsInRow(9)) * 4 >= CountFreeMainCells()))
            return true;

        if (best.Row < 6 && CountDice(best.Row + 1) == 3)
            return true;

        if (best.Row == 13 || best.Row == 12 && best.Score > 0)
            return true;

        return best.Score >= threshold;
    }

    private bool ShouldStopOnCompletedTacticalPrizeMove(int row)
    {
        if (row is < 0 or >= RowCount - 1)
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        var closesOwnPrize = CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column);
        var blocksOpponentPrize = ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);

        if (closesOwnPrize && blocksOpponentPrize)
            return true;

        return GetTacticalPrizeUrgency(row) >= 180;
    }

    private bool ShouldStopOnOpeningTwoPairs(Dictionary<int, int> counts, int score)
    {
        if (_rollCount != 1 || score < 18 || CurrentPlayer.GetFreeCell(7) == -1)
            return false;

        if (CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5))
            return false;
        if (CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4))
            return false;
        if (CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts))
            return false;
        if (ShouldDeferOpeningTwoPairsForSchoolDevelopment(counts))
            return false;
        if (!IsTacticalPrizeMove(7)
            && score <= 20
            && (CurrentPlayer.GetFreeCell(7) >= ColumnCount - 2 || CountFreeMainCells() <= 24))
            return false;

        var affordableSchoolRow = GetAffordableNegativeSchoolRow();
        if (!IsTacticalPrizeMove(7)
            && affordableSchoolRow >= 0
            && CalculateScore(affordableSchoolRow) >= -2)
            return false;

        return !HasHeavySchoolEndgamePressure();
    }

    // Главный выбор строки для записи результата.
    // Сначала идут жесткие тактические правила, которые должны перебить обычные веса:
    // сильная готовая комбинация, дешевый минус в школе, пара вместо слабых двух пар, школа после третьего броска.
    // Если жесткого правила нет, все свободные строки оцениваются общей функцией приоритета.
    private ComputerMove GetBestComputerMove()
    {
        var forcedPremiumRow = GetForcedPremiumMove();
        if (forcedPremiumRow >= 0 && IsLegalComputerMove(forcedPremiumRow, CalculateScore(forcedPremiumRow)))
            return new ComputerMove(forcedPremiumRow, CalculateScore(forcedPremiumRow), int.MaxValue);

        var forcedTacticalPrizeRow = GetForcedTacticalPrizeMove();
        if (forcedTacticalPrizeRow >= 0 && IsLegalComputerMove(forcedTacticalPrizeRow, CalculateScore(forcedTacticalPrizeRow)))
            return new ComputerMove(forcedTacticalPrizeRow, CalculateScore(forcedTacticalPrizeRow), int.MaxValue);

        var forcedSchoolRow = GetForcedSchoolMove();
        if (forcedSchoolRow >= 0 && IsLegalComputerMove(forcedSchoolRow, CalculateScore(forcedSchoolRow)))
            return new ComputerMove(forcedSchoolRow, CalculateScore(forcedSchoolRow), int.MaxValue);

        var forcedLowCostSchoolRow = GetForcedLowCostSchoolMove();
        if (forcedLowCostSchoolRow >= 0 && IsLegalComputerMove(forcedLowCostSchoolRow, CalculateScore(forcedLowCostSchoolRow)))
            return new ComputerMove(forcedLowCostSchoolRow, CalculateScore(forcedLowCostSchoolRow), int.MaxValue);

        var forcedTripleRow = GetForcedTripleMove();
        if (forcedTripleRow >= 0 && IsLegalComputerMove(forcedTripleRow, CalculateScore(forcedTripleRow)))
            return new ComputerMove(forcedTripleRow, CalculateScore(forcedTripleRow), int.MaxValue);

        var forcedUsefulPairRow = GetForcedUsefulPairMove();
        if (forcedUsefulPairRow >= 0 && IsLegalComputerMove(forcedUsefulPairRow, CalculateScore(forcedUsefulPairRow)))
            return new ComputerMove(forcedUsefulPairRow, CalculateScore(forcedUsefulPairRow), int.MaxValue);

        var forcedWeakResultCrossOutRow = GetForcedWeakResultCrossOutMove();
        if (forcedWeakResultCrossOutRow >= 0 && IsLegalComputerMove(forcedWeakResultCrossOutRow, CalculateScore(forcedWeakResultCrossOutRow)))
            return new ComputerMove(forcedWeakResultCrossOutRow, CalculateScore(forcedWeakResultCrossOutRow), int.MaxValue);

        var forcedWeakPairRow = GetForcedWeakPairMove();
        if (forcedWeakPairRow >= 0 && IsLegalComputerMove(forcedWeakPairRow, CalculateScore(forcedWeakPairRow)))
            return new ComputerMove(forcedWeakPairRow, CalculateScore(forcedWeakPairRow), int.MaxValue);

        var forcedDeadEndCrossOutRow = GetForcedDeadEndCrossOutMove();
        if (forcedDeadEndCrossOutRow >= 0 && IsLegalComputerMove(forcedDeadEndCrossOutRow, CalculateScore(forcedDeadEndCrossOutRow)))
            return new ComputerMove(forcedDeadEndCrossOutRow, CalculateScore(forcedDeadEndCrossOutRow), int.MaxValue);

        var bestRow = -1;
        var bestAdjusted = int.MinValue;
        var preferredAffordableSchoolRow = GetAffordableNegativeSchoolRow();

        // Общая оценка всех доступных строк: реальные очки хода плюс бонусы/штрафы за состояние таблицы.
        // Здесь же отсекаются запрещенные школьные минусы и слишком дорогие списания школы.
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
            if (IsProtectedLastColumnCrossOut(row, score))
                continue;
            if (IsProtectedPlayableCombinationCrossOut(row, score))
                continue;
            if (row < 6 && score < 0 && preferredAffordableSchoolRow >= 0 && row != preferredAffordableSchoolRow)
                continue;
            if (row < 6 && CountDice(row + 1) == 0 && HasFreeCombinationMainCell())
                continue;
            if (row < 6 && score < 0 && -score > CurrentPlayer.School && !CurrentPlayer.IsOnlySchoolFree())
                continue;
            if (row < 6 && score < 0 && HasFreeCombinationMainCell() && !ShouldAllowNegativeSchool(row, score))
                continue;
            if (row == 12 && score > 0 && !IsKareWorthRecording(GetDiceCounts()))
                continue;
            if (row == 9 && score > 0 && !IsFullHouseWorthRecording(GetDiceCounts()))
                continue;

            var adjusted = GetComputerMovePriority(row, score);
            if (row >= 6 && score == 0)
                adjusted -= GetComputerCrossOutPenalty(row);
            if (row < 6 && score < 0)
                adjusted -= ShouldAllowNegativeSchool(row, score)
                    ? 22 + row * 3
                    : CurrentPlayer.Difficulty == ComputerDifficulty.Careful ? 70 : 55;

            if (adjusted > bestAdjusted)
            {
                bestAdjusted = adjusted;
                bestRow = row;
            }
        }

        if (bestRow == -1)
        {
            foreach (var row in Enumerable.Range(0, RowCount - 1).Where(row => CurrentPlayer.GetFreeCell(row) != -1))
            {
                if (!IsLegalComputerMove(row, CalculateScore(row)))
                    continue;

                bestRow = row;
                break;
            }
        }

        if (bestRow == -1)
            bestRow = Enumerable.Range(0, RowCount - 1).First(row => CurrentPlayer.GetFreeCell(row) != -1);

        return new ComputerMove(bestRow, CalculateScore(bestRow), bestAdjusted);
    }

    // Жесткие решения для готовых сильных ходов.
    // Этот блок срабатывает до общей оценки весов, чтобы ИИ не испортил очевидно сильную раздачу
    // и мог тактически закрыть/заблокировать приз даже не самым большим количеством очков.
    private int GetForcedPremiumMove()
    {
        var counts = GetDiceCounts();

        if (CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5))
            return 13;

        var forcedOpeningMajorRow = GetForcedOpeningReadyMajorMove(counts);
        if (forcedOpeningMajorRow >= 0)
            return forcedOpeningMajorRow;

        if (ShouldPreferUnderfilledKareOverSchool(counts))
            return 12;

        if (ShouldPreferKareOverWeakTriple(counts))
            return 12;

        if (ShouldPreferKareOverWeakTwoPairs(counts))
            return 12;

        if (ShouldPreferUnderfilledFullHouseOverTriple(counts))
            return 9;

        var tacticalCombinationRow = GetForcedTacticalPrizeMove();
        if (tacticalCombinationRow >= 6)
            return tacticalCombinationRow;

        var balancedFinishedRow = GetBalancedFinishedMove(counts);
        if (balancedFinishedRow >= 0)
            return balancedFinishedRow;

        var balancedSchoolRow = GetBalancedReadySchoolMove(counts);
        if (balancedSchoolRow >= 0)
            return balancedSchoolRow;

        if (CurrentPlayer.GetFreeCell(12) != -1 && IsKareWorthRecording(counts) && (_rollCount >= 3 || !CanChaseAbaca()))
            return 12;

        if (CurrentPlayer.GetFreeCell(12) != -1 && CurrentPlayer.School >= 6 && counts.Any(pair => pair.Value >= 4 && pair.Key >= 5))
            return 12;

        if (_rollCount >= 2 && CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts))
            return 9;

        if (CurrentPlayer.GetFreeCell(11) != -1 && CalculateScore(11) > 0)
            return 11;

        if (CurrentPlayer.GetFreeCell(10) != -1 && CalculateScore(10) > 0)
            return 10;

        if (_rollCount != 1)
            return -1;

        if (CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4))
            return 12;

        if (CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts))
            return 9;

        if (ShouldPreferTwoPairsOverTripleWhenKareUnavailable(counts))
            return 7;

        if (CurrentPlayer.GetFreeCell(7) != -1
            && TwoPairsScore(counts) > 0
            && IsTacticalPrizeMove(7)
            && !ShouldDeferOpeningTwoPairsForOpenOnes(counts))
            return 7;

        if (ShouldTakeOpeningSum(counts))
            return 14;

        if (CurrentPlayer.GetFreeCell(8) != -1
            && counts.Any(pair => pair.Value >= 3)
            && (IsTacticalPrizeMove(8) || IsStrongOpeningTriple(counts)))
            return 8;

        if (CurrentPlayer.GetFreeCell(7) != -1
            && TwoPairsScore(counts) >= 18
            && !ShouldDeferOpeningTwoPairsForOpenOnes(counts))
            return 7;

        if (CurrentPlayer.GetFreeCell(6) != -1 && PairScore(counts) >= 10 && ShouldTakeOpeningPair())
            return 6;

        if (CurrentPlayer.GetFreeCell(7) != -1
            && HasPremiumTwoPairs(counts)
            && !ShouldDeferOpeningTwoPairsForOpenOnes(counts))
            return 7;

        if (CurrentPlayer.GetFreeCell(14) != -1 && IsGoodOpeningSum() && OpponentCanClaimRowPrizeSoon(14))
            return 14;

        if (CurrentPlayer.GetFreeCell(14) != -1 && IsGoodOpeningSum())
            return 14;

        return -1;
    }

    private bool ShouldPreferTwoPairsOverTripleWhenKareUnavailable(Dictionary<int, int> counts)
    {
        if (CurrentPlayer.GetFreeCell(7) == -1 || CurrentPlayer.GetFreeCell(8) == -1)
            return false;

        if (CurrentPlayer.GetFreeCell(12) != -1)
            return false;

        var fourKindValue = counts
            .Where(pair => pair.Value >= 4)
            .Select(pair => pair.Key)
            .DefaultIfEmpty(0)
            .Max();

        if (fourKindValue == 0)
            return false;

        if (CurrentPlayer.GetFreeCell(13) != -1
            && fourKindValue <= 2
            && CurrentPlayer.GetFreeCell(fourKindValue - 1) != -1)
            return false;

        var twoPairsScore = TwoPairsScore(counts);
        var tripleScore = CalculateScore(8);
        if (twoPairsScore <= 0 || tripleScore <= 0)
            return false;

        if (IsTacticalPrizeMove(8) && !IsTacticalPrizeMove(7))
            return false;

        return twoPairsScore >= tripleScore
            || CountBusyMainCellsInRow(7) <= CountBusyMainCellsInRow(8);
    }

    private int GetForcedOpeningReadyMajorMove(Dictionary<int, int> counts)
    {
        if (_rollCount != 1)
            return -1;

        if (CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4))
            return 12;

        if (CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts))
            return 9;

        if (CurrentPlayer.GetFreeCell(11) != -1 && CalculateScore(11) > 0)
            return 11;

        if (CurrentPlayer.GetFreeCell(10) != -1 && CalculateScore(10) > 0)
            return 10;

        if (CurrentPlayer.GetFreeCell(8) != -1
            && counts.Any(pair => pair.Value >= 3)
            && IsStrongOpeningTriple(counts)
            && !HasStraightFork()
            && !ShouldTakeOpeningSum(counts))
            return 8;

        return -1;
    }

    // Проверка тактической ценности хода: свой приз или блокировка приза соперника
    // учитываются и по строке таблицы, и по столбцу таблицы.
    private bool IsTacticalPrizeMove(int row)
    {
        var column = CurrentPlayer.GetFreeCell(row);
        return CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column)
            || ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);
    }

    private int GetForcedTacticalPrizeMove()
    {
        var counts = GetDiceCounts();
        return Enumerable.Range(0, RowCount - 1)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1 && IsTacticalPrizeMove(row))
            .Where(row => !ShouldDeferWeakBlockingTacticalMove(row, counts))
            .Select(row => (Row: row, Score: CalculateScore(row)))
            .Where(item => IsPlayableTacticalPrizeScore(item.Row, item.Score))
            .OrderByDescending(item => GetTacticalPrizeUrgency(item.Row))
            .ThenByDescending(item => item.Score)
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    private bool ShouldDeferWeakBlockingTacticalMove(int row, Dictionary<int, int> counts)
    {
        if (row is not (6 or 7 or 8 or 14))
            return false;

        if (row is 6 or 7 && HasStraightFork() && GetTacticalPrizeUrgency(row) < 140)
            return true;

        var column = CurrentPlayer.GetFreeCell(row);
        var takesOwnPrize = CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column);
        if (takesOwnPrize)
            return false;

        var blocksOpponent = ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);
        if (!blocksOpponent)
            return false;

        var ownRowBusy = CountBusyMainCellsInRow(row);
        var ownColumnBusy = column >= 0 && column < ColumnCount - 1
            ? CountBusyMainCellsInColumn(column)
            : 0;
        if (ownRowBusy >= ColumnCount - 2 || ownColumnBusy >= RowCount - 2)
            return false;
        if (GetTacticalPrizeUrgency(row) < 140 && HasStrongReadyAlternativeForWeakBlock(counts))
            return true;

        return CurrentPlayer.GetFreeCell(12) != -1 && IsKareWorthRecording(counts)
            || CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts)
            || row is 6 or 7 && CurrentPlayer.GetFreeCell(8) != -1 && counts.Any(pair => pair.Value >= 3)
            || counts.Any(pair => pair.Value >= 3 && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1);
    }

    private bool HasStrongReadyAlternativeForWeakBlock(Dictionary<int, int> counts)
    {
        if (CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5))
            return true;
        if (CurrentPlayer.GetFreeCell(12) != -1 && IsKareWorthRecording(counts))
            return true;
        if (CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts))
            return true;
        if (CurrentPlayer.GetFreeCell(11) != -1 && CalculateScore(11) > 0)
            return true;
        if (CurrentPlayer.GetFreeCell(10) != -1 && CalculateScore(10) > 0)
            return true;
        if (CurrentPlayer.GetFreeCell(8) != -1 && counts.Any(pair => pair.Value >= 3))
            return true;
        if (counts.Any(pair => pair.Value >= 2
            && pair.Key >= 3
            && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1))
            return true;

        return Enumerable.Range(0, 6)
            .Any(schoolRow => CurrentPlayer.GetFreeCell(schoolRow) != -1
                && CountDice(schoolRow + 1) >= 3);
    }

    private bool IsPlayableTacticalPrizeScore(int row, int score)
    {
        if (row < 6)
        {
            if (CountDice(row + 1) == 0)
                return false;
            if (score >= 0)
                return true;
            if (!ShouldAllowNegativeSchool(row, score))
                return false;

            return score >= -3 || !IsOnlyOpponentColumnBlock(row);
        }

        if (row == 14)
            return _rollCount == 1 && IsGoodOpeningSum()
                || CompletesOwnRowPrize(row)
                || CurrentPlayer.GetFreeCell(row) is >= 0 and < ColumnCount - 1 && CompletesOwnColumnPrize(CurrentPlayer.GetFreeCell(row));

        return score > 0;
    }

    private bool IsOnlyOpponentColumnBlock(int row)
    {
        var column = CurrentPlayer.GetFreeCell(row);
        if (column < 0 || column >= ColumnCount - 1)
            return false;

        return !CompletesOwnRowPrize(row)
            && !CompletesOwnColumnPrize(column)
            && !ShouldBlockOpponentRowPrize(row)
            && ShouldBlockOpponentColumnPrize(column);
    }

    private bool IsProtectedLastColumnCrossOut(int row, int score)
    {
        if (row < 6 || score != 0)
            return false;

        return CurrentPlayer.GetFreeCell(row) == ColumnCount - 2
            && HasEarlierCombinationCrossOutAvailable();
    }

    private bool HasEarlierCombinationCrossOutAvailable()
    {
        return Enumerable.Range(6, RowCount - 7)
            .Any(row =>
            {
                var column = CurrentPlayer.GetFreeCell(row);
                return column >= 0
                    && column < ColumnCount - 2
                    && CalculateScore(row) == 0;
            });
    }

    private bool IsProtectedPlayableCombinationCrossOut(int row, int score)
    {
        if (score != 0 || row is not (6 or 7 or 8 or 9 or 12))
            return false;

        return HasEarlierHardCombinationCrossOutAvailable();
    }

    private bool HasEarlierHardCombinationCrossOutAvailable()
    {
        return new[] { 13, 11, 10 }
            .Any(row =>
            {
                var column = CurrentPlayer.GetFreeCell(row);
                return column >= 0
                    && column < ColumnCount - 2
                    && CalculateScore(row) == 0;
            });
    }

    private bool ShouldPreferKareOverWeakTriple(Dictionary<int, int> counts)
    {
        if (_rollCount < 3 || CurrentPlayer.GetFreeCell(12) == -1 || CurrentPlayer.GetFreeCell(8) == -1)
            return false;
        if (!IsKareWorthRecording(counts))
            return false;

        var kareValue = GetKareValue(counts);
        if (kareValue == 0 || CountDice(kareValue) < 4)
            return false;

        var tripleScore = CalculateScore(8);
        if (tripleScore <= 0 || IsTacticalPrizeMove(8))
            return false;

        var kareRowFill = CountBusyMainCellsInRow(12);
        var tripleRowFill = CountBusyMainCellsInRow(8);
        var kareUrgency = GetTacticalPrizeUrgency(12);
        var tripleUrgency = GetTacticalPrizeUrgency(8);
        if (kareValue >= 5)
            return true;
        if (kareUrgency >= tripleUrgency + 10)
            return true;

        return tripleScore <= 6 && kareRowFill <= tripleRowFill
            || tripleScore <= 12 && kareRowFill + 1 <= tripleRowFill && !CompletesOwnRowPrize(8)
            || tripleScore <= 12 && kareRowFill + 2 <= tripleRowFill
            || CurrentPlayer.School >= 8 && kareRowFill == 0 && tripleRowFill >= 2;
    }

    private bool ShouldPreferKareOverWeakTwoPairs(Dictionary<int, int> counts)
    {
        if (_rollCount < 3 || CurrentPlayer.GetFreeCell(12) == -1 || CurrentPlayer.GetFreeCell(7) == -1)
            return false;
        if (!IsKareWorthRecording(counts))
            return false;

        var kareValue = GetKareValue(counts);
        if (kareValue == 0 || CountDice(kareValue) < 4)
            return false;

        var twoPairsScore = CalculateScore(7);
        if (twoPairsScore <= 0 || IsTacticalPrizeMove(7))
            return false;

        var kareUrgency = GetTacticalPrizeUrgency(12);
        var twoPairsUrgency = GetTacticalPrizeUrgency(7);
        if (kareUrgency >= twoPairsUrgency + 10)
            return true;

        return twoPairsScore <= 12;
    }

    private bool ShouldPreferUnderfilledKareOverSchool(Dictionary<int, int> counts)
    {
        if (_rollCount < 3 || CurrentPlayer.GetFreeCell(12) == -1 || !IsKareWorthRecording(counts))
            return false;

        var kareValue = GetKareValue(counts);
        if (kareValue < 3 || CountDice(kareValue) < 4)
            return false;

        var schoolRow = kareValue - 1;
        if (CurrentPlayer.GetFreeCell(schoolRow) == -1 || CountDice(kareValue) < 3 || CurrentPlayer.School < 8)
            return false;

        var schoolIsUrgent = CompletesOwnRowPrize(schoolRow)
            || ShouldBlockOpponentRowPrize(schoolRow);
        var schoolColumn = CurrentPlayer.GetFreeCell(schoolRow);
        if (schoolColumn >= 0 && schoolColumn < ColumnCount - 1)
        {
            schoolIsUrgent = schoolIsUrgent
                || CompletesOwnColumnPrize(schoolColumn)
                || ShouldBlockOpponentColumnPrize(schoolColumn);
        }

        if (schoolIsUrgent)
            return false;

        var kareRowFill = CountBusyMainCellsInRow(12);
        var schoolRowFill = CountBusyMainCellsInRow(schoolRow);
        if (kareValue >= 5)
            return kareRowFill <= 1
                || kareRowFill + 2 <= schoolRowFill;

        if (CurrentPlayer.School < 10)
            return false;

        return kareRowFill == 0 && schoolRowFill >= 3
            || kareRowFill + 3 <= schoolRowFill;
    }

    private bool ShouldPreferUnderfilledFullHouseOverTriple(Dictionary<int, int> counts)
    {
        if (_rollCount < 2 || CurrentPlayer.GetFreeCell(9) == -1 || CurrentPlayer.GetFreeCell(8) == -1)
            return false;

        var fullHouseScore = FullHouseScore(counts);
        var tripleScore = CalculateScore(8);
        if (fullHouseScore <= 0 || tripleScore <= 0)
            return false;

        if (ShouldBlockOpponentRowPrize(8))
            return false;

        var tripleColumn = CurrentPlayer.GetFreeCell(8);
        if (tripleColumn >= 0 && tripleColumn < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(tripleColumn))
            return false;

        var fullRowFill = CountBusyMainCellsInRow(9);
        var tripleRowFill = CountBusyMainCellsInRow(8);
        if (fullRowFill > 2)
            return false;

        return fullRowFill + 2 <= tripleRowFill
            || fullRowFill <= 1 && tripleScore <= 12;
    }

    private int GetTacticalPrizeUrgency(int row)
    {
        var column = CurrentPlayer.GetFreeCell(row);
        var urgency = 0;
        if (CompletesOwnRowPrize(row))
            urgency += ShouldBlockOpponentRowPrize(row) ? 140 : 90;
        if (ShouldBlockOpponentRowPrize(row))
            urgency += 120;
        if (column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column))
            urgency += ShouldBlockOpponentColumnPrize(column) ? 100 : 60;
        if (column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column))
            urgency += 80;

        return urgency;
    }

    // Общий выбор готового хода после последнего броска.
    // Он не отдает школу или каре приказом: школа, пара, две пары, тройка, фул, стриты, каре и абак
    // сравниваются по очкам, тактике приза и балансу заполнения строк/столбцов.
    private int GetBalancedFinishedMove(Dictionary<int, int> counts)
    {
        if (_rollCount < 3)
            return -1;

        return Enumerable.Range(0, 14)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1)
            .Select(row => (Row: row, Score: CalculateScore(row)))
            .Where(item => IsBalancedFinishedMoveCandidate(item.Row, item.Score, counts))
            .Select(item => new
            {
                item.Row,
                Adjusted = GetComputerMovePriority(item.Row, item.Score)
                    + GetFinishedMoveBalanceBonus(item.Row, item.Score)
            })
            .OrderByDescending(item => item.Adjusted)
            .ThenByDescending(item => CalculateScore(item.Row))
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    private bool IsBalancedFinishedMoveCandidate(int row, int score, Dictionary<int, int> counts)
    {
        if (row < 6)
            return score >= 0 && CountDice(row + 1) >= 3;

        if (score <= 0)
            return false;

        return row switch
        {
            6 => (score >= 8 || IsTacticalPrizeMove(row))
                && (IsTacticalPrizeMove(row) || !HasStrongerFinishedMoveThanPairRows(counts)),
            7 => (score >= 18 || IsTacticalPrizeMove(row))
                && (IsTacticalPrizeMove(row) || !HasStrongerFinishedMoveThanPairRows(counts)),
            8 => counts.Any(pair => pair.Value >= 3),
            9 => IsFullHouseWorthRecording(counts),
            10 or 11 => true,
            12 => IsKareWorthRecording(counts),
            13 => counts.Any(pair => pair.Value == 5),
            14 => IsTacticalPrizeMove(row) || score >= 26,
            _ => false
        };
    }

    private bool HasStrongerFinishedMoveThanPairRows(Dictionary<int, int> counts)
    {
        return CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5)
            || CurrentPlayer.GetFreeCell(12) != -1 && IsKareWorthRecording(counts)
            || CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts)
            || CurrentPlayer.GetFreeCell(11) != -1 && CalculateScore(11) > 0
            || CurrentPlayer.GetFreeCell(10) != -1 && CalculateScore(10) > 0
            || CurrentPlayer.GetFreeCell(8) != -1 && counts.Any(pair => pair.Value >= 3);
    }

    private int GetFinishedMoveBalanceBonus(int row, int score)
    {
        var bonus = 0;
        if (IsTacticalPrizeMove(row))
            bonus += 80;

        if (row < 6)
        {
            var rowBusy = CountBusyMainCellsInRow(row);
            var leastSchoolRow = Enumerable.Range(0, 6)
                .Where(schoolRow => CurrentPlayer.GetFreeCell(schoolRow) != -1)
                .Select(CountBusyMainCellsInRow)
                .DefaultIfEmpty(rowBusy)
                .Min();

            if (rowBusy <= leastSchoolRow + 1)
                bonus += 28;
            else if (rowBusy >= 4 && !CompletesOwnRowPrize(row))
                bonus -= 30;

            if (score == 0)
                bonus += 24;
            if (ShouldPreferCombinationsOverSchool() && !IsTacticalPrizeMove(row))
                bonus -= 70;
        }
        else
        {
            var rowBusy = CountBusyMainCellsInRow(row);
            var leastCombinationRow = Enumerable.Range(6, RowCount - 7)
                .Where(combinationRow => CurrentPlayer.GetFreeCell(combinationRow) != -1)
                .Select(CountBusyMainCellsInRow)
                .DefaultIfEmpty(rowBusy)
                .Min();

            bonus += score / 2;
            if (rowBusy <= leastCombinationRow + 1)
                bonus += 18;
            else if (rowBusy >= 4 && !CompletesOwnRowPrize(row))
                bonus -= 12;

            if (ShouldPreferCombinationsOverSchool())
                bonus += 30;
        }

        return bonus;
    }

    // Мягкая страховка для школы после третьего броска.
    // Школа берется только если она нужна по балансу/тактике или если нет более ценной готовой комбинации.
    private int GetForcedSchoolMove()
    {
        if (_rollCount < 3)
            return -1;

        var counts = GetDiceCounts();
        return counts
            .Where(pair => pair.Value >= 3 && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1)
            .Where(pair =>
            {
                var row = pair.Key - 1;
                if (ShouldDeferOverfilledLowSchoolToValuableCombination(row, counts))
                    return false;

                return ShouldPreferEndgameSchoolSecurity(row)
                    || ShouldPreferSchoolByTableBalance()
                    || IsTacticalPrizeMove(row)
                    || !HasValuableFinishedCombination(counts);
            })
            .OrderByDescending(pair => pair.Key)
            .Select(pair => pair.Key - 1)
            .FirstOrDefault(-1);
    }

    private bool ShouldDeferOverfilledLowSchoolToValuableCombination(int row, Dictionary<int, int> counts)
    {
        if (row is < 0 or > 2)
            return false;
        if (CountBusyMainCellsInRow(row) < 3 || CurrentPlayer.School < 8)
            return false;

        return CurrentPlayer.GetFreeCell(12) != -1 && IsKareWorthRecording(counts)
            || CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts);
    }

    private bool ShouldPreferEndgameSchoolSecurity(int row)
    {
        if (row is < 0 or >= 6 || _rollCount < 3 || CountFreeMainCells() > 8)
            return false;
        if (CurrentPlayer.GetFreeCell(row) == -1 || CountDice(row + 1) < 3)
            return false;

        var value = row + 1;
        return value >= 5 && CurrentPlayer.School < value * 2
            || value >= 4 && CountDice(value) >= 4 && HasHeavySchoolEndgamePressure()
            || HasComfortableLeadForSchoolLock() && HasHeavySchoolEndgamePressure();
    }

    // Если последним броском получилась только слабая комбинация,
    // маленький допустимый минус в школе может быть лучше, чем тратить важную строку комбинаций.
    private int GetForcedLowCostSchoolMove()
    {
        if (_rollCount < 3 || CurrentPlayer.School <= 0)
            return -1;

        var finalSchoolCellRow = GetLastOpenSchoolCellRow();
        if (finalSchoolCellRow >= 0 && CountDice(finalSchoolCellRow + 1) > 0)
            return finalSchoolCellRow;

        var safeSchoolReserveRow = GetSafeSchoolReserveCloseRow();
        if (safeSchoolReserveRow >= 0)
            return safeSchoolReserveRow;

        var affordableSchoolRow = GetAffordableNegativeSchoolRow();
        if (affordableSchoolRow < 0)
            return -1;

        var affordableSchoolScore = CalculateScore(affordableSchoolRow);
        if (affordableSchoolScore == -1
            && !IsTacticalPrizeMove(affordableSchoolRow)
            && !CompletesOwnRowPrize(affordableSchoolRow)
            && !ShouldBlockOpponentRowPrize(affordableSchoolRow))
            return affordableSchoolRow;

        if (ShouldForceCheapSchoolOverCrossOut(affordableSchoolRow))
            return affordableSchoolRow;
        if (ShouldPreferSacrificialCrossOutOverSchool(affordableSchoolRow))
            return -1;

        if (affordableSchoolScore >= -1 && !IsTacticalPrizeMove(affordableSchoolRow))
            return affordableSchoolRow;

        var bestWeakCombination = GetBestWeakCombinationMove();
        if (bestWeakCombination.Row >= 0
            && (bestWeakCombination.Score > 14 || IsTacticalPrizeMove(bestWeakCombination.Row))
            && !ShouldPreferAffordableSchoolOverWeakCombination(bestWeakCombination))
            return -1;

        return affordableSchoolRow;
    }

    private int GetLastOpenSchoolCellRow()
    {
        var freeSchoolCells = 0;
        var lastRow = -1;

        for (var row = 0; row < 6; row++)
        {
            for (var column = 0; column < ColumnCount - 1; column++)
            {
                if (CurrentPlayer.Table[row, column] != EmptyCell)
                    continue;

                freeSchoolCells++;
                lastRow = row;
                if (freeSchoolCells > 1)
                    return -1;
            }
        }

        return freeSchoolCells == 1 ? lastRow : -1;
    }

    private int GetSafeSchoolReserveCloseRow()
    {
        return Enumerable.Range(0, 6)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1 && CountDice(row + 1) > 0)
            .Select(row => (Row: row, Score: CalculateScore(row)))
            .Where(item => item.Score < 0 && CanSafelyCloseRemainingSchool(item.Row, item.Score))
            .OrderByDescending(item => item.Score)
            .ThenByDescending(item => CountBusyMainCellsInRow(item.Row))
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    private bool ShouldPreferAffordableSchoolOverWeakCombination((int Row, int Score) move)
    {
        return move.Row switch
        {
            6 => move.Score <= 12,
            7 => move.Score < 18,
            14 => move.Score <= 26 && !OpponentCanClaimRowPrizeSoon(14),
            _ => move.Score <= 6
        };
    }

    private bool ShouldPreferSacrificialCrossOutOverSchool(int schoolRow)
    {
        var score = CalculateScore(schoolRow);
        if (score >= -2)
            return false;

        var sacrificialRow = GetSacrificialCrossOutRow();
        if (sacrificialRow < 0 || ShouldProtectOwnPrizeProgressBeforeCrossOut(sacrificialRow))
            return false;

        var selfDamage = GetSchoolNegativeSelfDamage(schoolRow);
        var remainingSchool = CurrentPlayer.School + score;
        return score <= -4
            && HasSacrificialCrossOutAvailable()
            && !CompletesOwnRowPrize(schoolRow)
            && !ShouldBlockOpponentRowPrize(schoolRow)
            && (selfDamage >= 40 || CountBusyMainCellsInRow(schoolRow) >= 3 || remainingSchool <= 3)
            && !(score >= -3 && remainingSchool >= 4);
    }

    private bool ShouldForceCheapSchoolOverCrossOut(int schoolRow)
    {
        if (schoolRow is < 0 or >= 6 || CurrentPlayer.GetFreeCell(schoolRow) == -1)
            return false;

        var score = CalculateScore(schoolRow);
        if (score > -2 || !IsLegalComputerMove(schoolRow, score))
            return false;

        return !IsTacticalPrizeMove(schoolRow)
            && !CompletesOwnRowPrize(schoolRow)
            && !ShouldBlockOpponentRowPrize(schoolRow);
    }

    private bool ShouldProtectOwnPrizeProgressBeforeCrossOut(int row)
    {
        if (row < 6 || CurrentPlayer.GetFreeCell(row) == -1)
            return false;

        var rowBusy = CountBusyMainCellsInRow(row);
        if (rowBusy >= ColumnCount - 3)
            return true;

        var column = CurrentPlayer.GetFreeCell(row);
        return column >= 0
            && column < ColumnCount - 1
            && CountBusyMainCellsInColumn(column) >= RowCount - 3;
    }

    // Для слабых двух пар до 18 очков ИИ сначала пробует школу,
    // а если школа недоступна, выбирает пару, чтобы сохранить строку "две пары" для более сильного хода.
    private int GetForcedWeakPairMove()
    {
        if (_rollCount < 3 || CurrentPlayer.GetFreeCell(6) == -1)
            return -1;

        if (HasBetterFinishedMoveThanWeakTwoPairs())
            return -1;

        var twoPairsScore = CalculateScore(7);
        if (twoPairsScore <= 0 || IsTacticalPrizeMove(7))
            return -1;

        if (twoPairsScore > 18 || twoPairsScore == 18 && !ShouldPreferPairOverMediumTwoPairs())
            return -1;

        return CalculateScore(6) > 0 ? 6 : -1;
    }

    private int GetForcedDeadEndCrossOutMove()
    {
        if (_rollCount < 3 || HasPlayableNonCrossOutMove())
            return -1;

        return GetEmergencySacrificialCrossOutRow();
    }

    private bool HasPlayableNonCrossOutMove()
    {
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
            if (row < 6)
            {
                if (CountDice(row + 1) > 0 && score >= 0)
                    return true;
                if (score < 0 && ShouldAllowNegativeSchool(row, score))
                    return true;

                continue;
            }

            if (score > 0)
                return true;
        }

        return false;
    }

    private bool ShouldPreferPairOverMediumTwoPairs()
    {
        return CountBusyMainCellsInRow(6) == 0 && CountBusyMainCellsInRow(7) > 0;
    }

    private int PairScore(Dictionary<int, int> counts)
    {
        return counts
            .Where(pair => pair.Value >= 2)
            .Select(pair => pair.Key * 2)
            .DefaultIfEmpty(0)
            .Max();
    }

    private int GetKareValue(Dictionary<int, int> counts)
    {
        return counts
            .Where(pair => pair.Value >= 4)
            .Select(pair => pair.Key)
            .DefaultIfEmpty(0)
            .Max();
    }

    private bool IsKareWorthRecording(Dictionary<int, int> counts)
    {
        var value = GetKareValue(counts);
        if (value == 0 || CurrentPlayer.GetFreeCell(12) == -1)
            return false;

        if (_rollCount == 1)
            return true;

        if (value >= 5)
            return true;

        var kareRowFill = CountBusyMainCellsInRow(12);
        if (_rollCount >= 3 && CurrentPlayer.GetFreeCell(value - 1) == -1)
            return true;

        if (_rollCount >= 3 && value >= 2 && kareRowFill <= 1)
            return true;

        if (_rollCount >= 3 && value >= 3 && CurrentPlayer.GetFreeCell(value - 1) == -1 && kareRowFill <= 2)
            return true;

        var kareRowIsAlreadyUsed = kareRowFill >= 2;
        if (CurrentPlayer.School <= 6 || kareRowIsAlreadyUsed)
            return false;

        return value == 4 && IsTacticalPrizeMove(12);
    }

    private bool IsFullHouseWorthRecording(Dictionary<int, int> counts)
    {
        var score = FullHouseScore(counts);
        if (score == 0 || CurrentPlayer.GetFreeCell(9) == -1)
            return false;

        if (_rollCount == 1)
            return true;

        if (IsTacticalPrizeMove(9))
            return true;

        if (score >= 21)
            return true;

        var fullRowFill = CountBusyMainCellsInRow(9);
        if (_rollCount >= 2 && fullRowFill >= 3 && score < 18)
            return false;

        var fullRowIsAlreadyUsed = fullRowFill >= 2;
        if (_rollCount >= 3 && score >= 18 && fullRowFill <= 2)
            return true;

        if (_rollCount >= 3 && score > 0 && HasSacrificialCrossOutAvailable() && fullRowFill <= 1)
            return true;

        if (CurrentPlayer.School <= 6 || fullRowIsAlreadyUsed)
            return false;

        return false;
    }

    private bool ShouldTakeOpeningPair()
    {
        return _rollCount == 1
            && CountBusyMainCellsInRow(6) == 0
            && !GetDiceCounts().Any(pair => pair.Value >= 3)
            && !GetDiceCounts().Any(pair => pair.Key >= 3 && pair.Value >= 2 && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1)
            && ShouldPreferCombinationsOverSchool()
            && IsTacticalPrizeMove(6);
    }

    private int GetForcedTripleMove()
    {
        if (CurrentPlayer.GetFreeCell(8) == -1 || !GetDiceCounts().Any(pair => pair.Value >= 3))
            return -1;

        var hasStrongerSameDiceMove =
            CurrentPlayer.GetFreeCell(13) != -1 && CalculateScore(13) > 0
            || CurrentPlayer.GetFreeCell(12) != -1 && CalculateScore(12) > 0
            || CurrentPlayer.GetFreeCell(9) != -1 && CalculateScore(9) > 0;

        return hasStrongerSameDiceMove ? -1 : 8;
    }

    private int GetForcedUsefulPairMove()
    {
        if (_rollCount < 3 || CurrentPlayer.GetFreeCell(6) == -1)
            return -1;

        var pairScore = CalculateScore(6);
        if (pairScore < 6 || HasBetterFinishedMoveThanWeakTwoPairs())
            return -1;

        var affordableSchoolRow = GetAffordableNegativeSchoolRow();
        if (affordableSchoolRow >= 0
            && !IsTacticalPrizeMove(affordableSchoolRow)
            && CalculateScore(affordableSchoolRow) >= -1
            && CurrentPlayer.School >= 8
            && pairScore <= 14)
            return -1;

        return 6;
    }

    private int GetForcedWeakResultCrossOutMove()
    {
        if (_rollCount < 3)
            return -1;

        var affordableSchoolRow = GetAffordableNegativeSchoolRow();
        if (affordableSchoolRow >= 0 && !ShouldPreferSacrificialCrossOutOverSchool(affordableSchoolRow))
            return -1;

        var sacrificialRow = GetSacrificialCrossOutRow();
        if (sacrificialRow < 0)
            return -1;

        var pairScore = CurrentPlayer.GetFreeCell(6) != -1 && !IsTacticalPrizeMove(6)
            ? CalculateScore(6)
            : 0;
        var twoPairsScore = CurrentPlayer.GetFreeCell(7) != -1 && !IsTacticalPrizeMove(7)
            ? CalculateScore(7)
            : 0;
        var sumScore = CurrentPlayer.GetFreeCell(14) != -1 && !IsTacticalPrizeMove(14)
            ? CalculateScore(14)
            : 0;

        var hasOnlyWeakPairResult = pairScore is > 0 and <= 8;
        var hasOnlyWeakTwoPairsResult = twoPairsScore is > 0 and < 18;
        var hasOnlyWeakSumResult = sumScore is > 0 and <= 22;
        var bestUsefulCombinationScore = GetBestWeakCombinationMove().Score;
        if (bestUsefulCombinationScore > 0
            && !hasOnlyWeakPairResult
            && !hasOnlyWeakTwoPairsResult
            && !hasOnlyWeakSumResult)
            return -1;

        return sacrificialRow;
    }

    private int GetAffordableNegativeSchoolRow()
    {
        return Enumerable.Range(0, 6)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1)
            .Select(row => (Row: row, Score: CalculateScore(row), DiceCount: CountDice(row + 1)))
            .Where(item => item.DiceCount > 0
                && (item.Score == 0 || item.Score < 0 && ShouldAllowNegativeSchool(item.Row, item.Score)))
            .OrderByDescending(item => item.Score)
            .ThenBy(item => GetSchoolNegativeSelfDamage(item.Row))
            .ThenBy(item => CountBusyMainCellsInRow(item.Row))
            .ThenBy(item => item.Row)
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    private int GetSchoolNegativeSelfDamage(int row)
    {
        if (row is < 0 or >= 6 || CurrentPlayer.GetFreeCell(row) == -1)
            return int.MaxValue;

        var rowBusy = CountBusyMainCellsInRow(row);
        var column = CurrentPlayer.GetFreeCell(row);
        var columnBusy = column >= 0 && column < ColumnCount - 1
            ? CountBusyMainCellsInColumn(column)
            : 0;

        var damage = rowBusy * 10 + columnBusy * 6;
        if (CompletesOwnRowPrize(row))
            damage += 40;
        if (column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column))
            damage += 40;

        return damage;
    }

    private int GetSacrificialCrossOutRow()
    {
        return new[] { 10, 11, 13, 9 }
            .Select((row, index) => new
            {
                Row = row,
                Order = index,
                Column = CurrentPlayer.GetFreeCell(row),
                RowCrosses = CountCrossesInMainRow(row),
                RowBusy = CountBusyMainCellsInRow(row)
            })
            .Where(item => item.Column >= 0
                && item.Column < ColumnCount - 2
                && CalculateScore(item.Row) == 0)
            .Select(item => new
            {
                item.Row,
                Priority = CountCrossesInMainColumn(item.Column) * 40
                    - item.RowBusy * 32
                    - item.RowCrosses * 24
                    + GetSacrificialOrderBonus(item.Row)
                    - item.Order
            })
            .OrderByDescending(item => item.Priority)
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    private int GetEmergencySacrificialCrossOutRow()
    {
        var emergencyRows = new[] { 10, 11, 13, 9, 12, 8 };
        var hasEarlierEmergencyOption = emergencyRows
            .Any(row =>
            {
                var column = CurrentPlayer.GetFreeCell(row);
                return column >= 0
                    && column < ColumnCount - 2
                    && CalculateScore(row) == 0;
            });

        return emergencyRows
            .Select((row, index) => new
            {
                Row = row,
                Order = index,
                Column = CurrentPlayer.GetFreeCell(row),
                RowCrosses = CountCrossesInMainRow(row),
                RowBusy = CountBusyMainCellsInRow(row)
            })
            .Where(item => item.Column >= 0
                && item.Column < ColumnCount - 1
                && CalculateScore(item.Row) == 0)
            .Select(item => new
            {
                item.Row,
                Priority = (item.Column == ColumnCount - 2
                        ? hasEarlierEmergencyOption ? -220 : -60
                        : 0)
                    + CountCrossesInMainColumn(item.Column) * 28
                    - item.RowBusy * 28
                    - item.RowCrosses * 18
                    + GetSacrificialOrderBonus(item.Row)
                    - item.Order * 12
            })
            .OrderByDescending(item => item.Priority)
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    private static int GetSacrificialOrderBonus(int row)
    {
        return row switch
        {
            10 => 24,
            11 => 16,
            13 => 8,
            9 => 0,
            12 => -8,
            8 => -12,
            _ => 0
        };
    }

    private int GetBalancedReadySchoolMove(Dictionary<int, int> counts)
    {
        if (_rollCount < 3)
            return -1;

        var readySchoolRows = counts
            .Where(pair => pair.Value >= 3 && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1)
            .OrderByDescending(pair => pair.Value)
            .ThenByDescending(pair => pair.Key)
            .Select(pair => pair.Key - 1)
            .ToArray();

        if (readySchoolRows.Length == 0 || !ShouldPreferSchoolByTableBalance())
            return -1;

        return readySchoolRows[0];
    }

    private bool ShouldPreferSchoolByTableBalance()
    {
        const double schoolCapacity = 6.0 * (ColumnCount - 1);
        const double combinationCapacity = (RowCount - 7.0) * (ColumnCount - 1);

        var schoolFill = CountBusyMainCells(0, 6) / schoolCapacity;
        var combinationFill = CountBusyMainCells(6, RowCount - 1) / combinationCapacity;

        return CurrentPlayer.School <= 0 && schoolFill <= combinationFill + 0.08
            || schoolFill + 0.12 < combinationFill;
    }

    private bool HasBetterFinishedMoveThanWeakTwoPairs()
    {
        var counts = GetDiceCounts();

        return CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5)
            || CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4)
            || CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts)
            || CurrentPlayer.GetFreeCell(8) != -1 && counts.Any(pair => pair.Value >= 3)
            || Enumerable.Range(0, 6).Any(row => CurrentPlayer.GetFreeCell(row) != -1 && CountDice(row + 1) >= 3);
    }

    // Ищет лучшую слабую комбинацию, чтобы сравнить ее с мягким списанием школы.
    private (int Row, int Score) GetBestWeakCombinationMove()
    {
        var bestRow = -1;
        var bestScore = 0;

        for (var row = 6; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
            if (score > bestScore)
            {
                bestScore = score;
                bestRow = row;
            }
        }

        return (bestRow, bestScore);
    }

    // Защита от плохого принудительного хода в школу:
    // если уже есть ценная готовая комбинация, ее нужно рассматривать раньше школьной записи.
    private bool HasValuableFinishedCombination(Dictionary<int, int> counts)
    {
        if (CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5))
            return true;
        if (CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4))
            return true;
        if (CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts))
            return true;
        if (CurrentPlayer.GetFreeCell(7) != -1 && TwoPairsScore(counts) > 0)
            return true;
        if (CurrentPlayer.GetFreeCell(8) != -1 && counts.Any(pair => pair.Value == 3))
            return _rollCount == 1;

        return false;
    }

    private static bool HasPremiumTwoPairs(Dictionary<int, int> counts)
    {
        var pairValues = GetTwoPairValues(counts);

        return pairValues.Count(value => value == 6) >= 2
            || pairValues.Count(value => value == 5) >= 2
            || pairValues.Contains(6) && pairValues.Contains(5)
            || pairValues.Contains(6) && pairValues.Contains(4)
            || pairValues.Contains(5) && pairValues.Contains(4);
    }

    private static bool IsStrongOpeningTriple(Dictionary<int, int> counts)
    {
        return counts.Any(pair => pair.Value >= 3 && pair.Key >= 5);
    }

    private bool IsGoodOpeningSum()
    {
        return _rollCount == 1 && _dice.Sum() >= 23;
    }

    // Хорошая сумма с раздачи может быть выгоднее слабой пары/тройки.
    // Но она не перебивает уже готовые сильные комбинации: абак, каре, фул или стрит.
    private bool ShouldTakeOpeningSum(Dictionary<int, int> counts)
    {
        if (CurrentPlayer.GetFreeCell(14) == -1 || !IsGoodOpeningSum())
            return false;

        var hasStrongReadyCombination =
            CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5)
            || CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4)
            || CurrentPlayer.GetFreeCell(9) != -1 && IsFullHouseWorthRecording(counts)
            || CurrentPlayer.GetFreeCell(11) != -1 && CalculateScore(11) > 0
            || CurrentPlayer.GetFreeCell(10) != -1 && CalculateScore(10) > 0;

        return !hasStrongReadyCombination
            && ShouldPreferOpeningSumOverReadyTwoPairs(counts);
    }

    private bool ShouldPreferOpeningSumOverReadyTwoPairs(Dictionary<int, int> counts)
    {
        if (CurrentPlayer.GetFreeCell(7) == -1 || !HasPremiumTwoPairs(counts))
            return true;
        if (IsTacticalPrizeMove(7) && !IsTacticalPrizeMove(14))
            return false;
        if (IsTacticalPrizeMove(14) && !IsTacticalPrizeMove(7))
            return true;

        return CountBusyMainCellsInRow(14) <= CountBusyMainCellsInRow(7);
    }

    private bool ShouldDeferOpeningTwoPairsForOpenOnes(Dictionary<int, int> counts)
    {
        if (_rollCount != 1)
            return false;
        if (CurrentPlayer.GetFreeCell(0) == -1 || CountBusyMainCellsInRow(0) > 3)
            return false;
        if (CountDice(1) < 2)
            return false;
        if (HasPremiumTwoPairs(counts))
            return false;

        var twoPairsColumn = CurrentPlayer.GetFreeCell(7);
        return !CompletesOwnRowPrize(7)
            && !(twoPairsColumn >= 0 && twoPairsColumn < ColumnCount - 1 && CompletesOwnColumnPrize(twoPairsColumn))
            && !ShouldBlockOpponentRowPrize(7)
            && !(twoPairsColumn >= 0 && twoPairsColumn < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(twoPairsColumn));
    }

    private bool ShouldDeferOpeningTwoPairsForSchoolDevelopment(Dictionary<int, int> counts)
    {
        if (_rollCount != 1 || HasPremiumTwoPairs(counts))
            return false;

        return counts.Any(pair =>
            pair.Value >= 2
            && pair.Key >= 4
            && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1
            && CountBusyMainCellsInRow(pair.Key - 1) <= 3
            && !IsTacticalPrizeMove(pair.Key - 1));
    }

    private bool CanChaseAbaca()
    {
        return CurrentPlayer.GetFreeCell(13) != -1 && GetDiceCounts().Any(pair => pair.Value == 4);
    }

    private bool CompletesOwnRowPrize(int row)
    {
        if (CurrentPlayer.Table[row, ColumnCount - 1] != EmptyCell)
            return false;

        var busyMainCells = 0;
        for (var column = 0; column < ColumnCount - 1; column++)
        {
            if (CurrentPlayer.Table[row, column] != EmptyCell)
                busyMainCells++;
        }

        return busyMainCells == ColumnCount - 2;
    }

    private bool CompletesOwnColumnPrize(int column)
    {
        if (CurrentPlayer.Table[RowCount - 1, column] != EmptyCell)
            return false;

        var busyRows = 0;
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.Table[row, column] != EmptyCell)
                busyRows++;
        }

        return busyRows == RowCount - 2;
    }

    private bool OpponentPrizeIsOpen(int row, int column)
    {
        return _players[1 - _currentPlayerIndex].Table[row, column] == EmptyCell;
    }

    private bool OpponentCanClaimRowPrizeSoon(int row)
    {
        var opponent = _players[1 - _currentPlayerIndex];
        if (opponent.Table[row, ColumnCount - 1] != EmptyCell)
            return false;

        var busyMainCells = 0;
        for (var column = 0; column < ColumnCount - 1; column++)
        {
            if (opponent.Table[row, column] != EmptyCell)
                busyMainCells++;
        }

        return busyMainCells == ColumnCount - 2;
    }

    private bool OpponentCanClaimColumnPrizeSoon(int column)
    {
        var opponent = _players[1 - _currentPlayerIndex];
        if (opponent.Table[RowCount - 1, column] != EmptyCell)
            return false;

        var busyRows = 0;
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (opponent.Table[row, column] != EmptyCell)
                busyRows++;
        }

        return busyRows == RowCount - 2;
    }

    private bool ShouldBlockOpponentRowPrize(int row)
    {
        if (!OpponentCanClaimRowPrizeSoon(row))
            return false;

        if (CompletesOwnRowPrize(row))
            return true;

        var threat = GetOpponentPrizeThreat(row);
        if (threat >= 70)
            return true;
        if (threat >= 45)
            return CountBusyMainCellsInRow(row) >= ColumnCount - 3;

        return false;
    }

    private bool ShouldBlockOpponentColumnPrize(int column)
    {
        if (!OpponentCanClaimColumnPrizeSoon(column))
            return false;

        if (CompletesOwnColumnPrize(column))
            return true;

        var missingRow = GetOpponentMissingRowForColumnPrize(column);
        var threat = GetOpponentPrizeThreat(missingRow);
        if (threat >= 70)
            return true;
        if (threat >= 45)
            return CountBusyMainCellsInColumn(column) >= RowCount - 3;

        return false;
    }

    private int GetOpponentMissingRowForColumnPrize(int column)
    {
        var opponent = _players[1 - _currentPlayerIndex];
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (opponent.Table[row, column] == EmptyCell)
                return row;
        }

        return -1;
    }

    private static int GetOpponentPrizeThreat(int row)
    {
        return row switch
        {
            >= 0 and < 6 => 78, // школа ловится часто, особенно если осталась одна строка до приза
            6 => 88,            // пара
            7 => 82,            // две пары
            8 => 74,            // тройка
            14 => 86,           // сумма
            9 => 54,            // фул
            12 => 50,           // каре
            10 or 11 => 28,     // стриты
            13 => 16,           // абак
            _ => 0
        };
    }

    // Выбор кубиков, которые ИИ оставляет перед следующим броском.
    // Это вторая половина ИИ: сначала GetBestComputerMove решает, куда хотелось бы записать ход,
    // а здесь компьютер фиксирует кости, которые помогают прийти к этой цели.
    private void ApplyComputerKeepStrategy()
    {
        Array.Fill(_fixedDice, false);
        Array.Fill(_selectedDice, false);

        var counts = Enumerable.Range(1, 6)
            .Select(value => (Value: value, Count: _dice.Count(die => die == value)))
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .ToArray();

        // Если все комбинации уже заняты/вычеркнуты, ИИ не должен ловить случайные пары или стриты.
        // В этом режиме фиксируются только кости номиналов, которые еще открыты в школе.
        if (!HasFreeCombinationMainCell() && TryKeepOnlyOpenSchoolDice())
            return;

        if (TryKeepStrongOnesBeforeUrgentTarget(counts))
            return;

        if (TryKeepOpeningTripleDevelopment(counts))
            return;

        if (TryKeepReadySchoolTripleBeforeUrgentTarget(counts))
            return;

        if (TryKeepCriticalEndgameSchoolDice(counts))
            return;

        if (TryKeepEndgameTargetDice())
            return;

        if (TryKeepOpenOnesBeforeUrgentTarget(counts))
            return;

        if (TryKeepOpenSchoolPairBeforeUrgentTarget(counts))
            return;

        // Если строка или столбец почти закрывают приз, сначала пытаемся удерживать кости под этот срочный ряд.
        var urgentTargetRow = GetUrgentComputerTargetRow();
        if (urgentTargetRow >= 0
            && !ShouldPreferStraightForkOverUrgentTarget(urgentTargetRow)
            && !ShouldPreferPairOverUrgentStraight(urgentTargetRow, counts)
            && ApplyComputerTargetKeepStrategy(urgentTargetRow, counts))
            return;

        if (TryKeepOpenSchoolSingleWhenOnlyStraightsRemain(counts))
            return;

        var smallStraight = Enumerable.Range(1, 5).Count(value => _dice.Contains(value));
        var bigStraight = Enumerable.Range(2, 5).Count(value => _dice.Contains(value));
        var straightTarget = CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive ? 3 : 4;
        // Стриты ловим только при открытых клетках стрита и подходящей заготовке.
        // Особый случай 2-3-4-5 дает шанс сразу на малый или большой стрит.
        if (HasOpenStraightTarget(smallStraight, bigStraight, straightTarget, counts))
        {
            if (HasStraightFork())
            {
                KeepStraightForkDice();
                return;
            }

            var straightValues = GetStraightKeepValues(smallStraight, bigStraight);
            KeepSingleDiceForValues(straightValues);
            return;
        }

        if (counts[0].Count >= 3)
        {
            KeepMostRepeated(counts);
            return;
        }

        if (TryKeepOpenOnesForSchool(counts))
            return;

        if (TryKeepOpenOnesOverClosedSchoolPair(counts))
            return;

        if (TryKeepUnderfilledFullHouseCandidate(counts))
            return;

        if (TryKeepCriticalSchoolSingleOverDeadPair(counts))
            return;

        if (TryKeepUsefulPairDice(counts))
            return;

        if (TryKeepHighSingleInsteadOfWeakLowPair(counts))
            return;

        // Обычная школьная стратегия: держим номинал школы только когда он действительно открыт и полезен.
        var schoolTarget = GetComputerSchoolTarget(counts);
        if (schoolTarget > 0)
        {
            for (var i = 0; i < DiceCount; i++)
                _fixedDice[i] = _dice[i] == schoolTarget;
            return;
        }

        if (CurrentPlayer.Difficulty == ComputerDifficulty.Careful)
        {
            var pair = counts.FirstOrDefault(item => item.Count >= 2);
            var pairValue = pair.Count >= 2 ? pair.Value : counts[0].Value;
            for (var i = 0; i < DiceCount; i++)
                _fixedDice[i] = _dice[i] == pairValue || (pair.Count < 2 && _rollCount >= 2 && _dice[i] >= 5);
            return;
        }

        var keepValue = counts[0].Value;
        if (CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive && counts[0].Count < 3)
            keepValue = counts.OrderByDescending(item => item.Value).First().Value;

        for (var i = 0; i < DiceCount; i++)
            _fixedDice[i] = _dice[i] == keepValue;
    }

    private bool TryKeepOpenSchoolSingleWhenOnlyStraightsRemain((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;

        var openCombinationRows = Enumerable.Range(6, RowCount - 7)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1)
            .ToArray();
        if (openCombinationRows.Length == 0 || openCombinationRows.Any(row => row is not (10 or 11)))
            return false;

        var schoolValue = counts
            .Where(item => item.Count > 0 && CurrentPlayer.GetFreeCell(item.Value - 1) != -1)
            .OrderByDescending(item => GetCriticalSchoolKeepPriority(item.Value, item.Count))
            .ThenByDescending(item => item.Value)
            .Select(item => item.Value)
            .FirstOrDefault();
        if (schoolValue == 0)
            return false;

        KeepValues([schoolValue]);
        return true;
    }

    // Выбирает номинал школы, который стоит ловить.
    // Если комбинации сейчас важнее школы, низкие или случайные школьные цели отбрасываются.
    private int GetComputerSchoolTarget(IEnumerable<(int Value, int Count)> counts)
    {
        var preferCombinations = ShouldPreferCombinationsOverSchool();
        var bestValue = 0;
        var bestPriority = int.MinValue;

        foreach (var item in counts)
        {
            var value = item.Value;
            var count = item.Count;
            var row = value - 1;
            if (!ShouldChaseSchoolValue(value, count, preferCombinations) || CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var rowFill = CountBusyMainCellsInRow(row);
            var priority = count * 12 + value + Math.Max(0, 4 - rowFill) * 10 + GetSchoolPrizeRaceBonus(row);
            if (CurrentPlayer.Difficulty == ComputerDifficulty.Careful)
                priority += value >= 5 ? 5 : 0;

            if (priority > bestPriority)
            {
                bestPriority = priority;
                bestValue = value;
            }
        }

        return bestValue;
    }

    // Режим конца партии: вне школы больше нет свободных клеток.
    // Если открыта только школа пятерок, фиксируем только пятерки; если пятерок нет, перебрасываем все.
    private bool TryKeepOnlyOpenSchoolDice()
    {
        var bestValue = 0;
        var bestCount = -1;
        for (var row = 0; row < 6; row++)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var value = row + 1;
            var count = CountDice(value);
            if (count > bestCount)
            {
                bestCount = count;
                bestValue = value;
            }
        }

        if (bestValue == 0)
            return false;

        for (var i = 0; i < DiceCount; i++)
            _fixedDice[i] = _dice[i] == bestValue && bestCount > 0;

        return true;
    }

    private bool TryKeepEndgameTargetDice()
    {
        if (CountFreeMainCells() > 4)
            return false;

        var bestStraightValues = GetBestEndgameStraightValues();
        var schoolValue = GetBestOpenSchoolValue();
        if (bestStraightValues.Count > 0)
        {
            var straightKept = bestStraightValues.Count(value => _dice.Contains(value));
            var schoolKept = schoolValue > 0 ? CountDice(schoolValue) : 0;
            if (straightKept >= Math.Max(3, schoolKept + 1))
            {
                KeepSingleDiceForValues(bestStraightValues);
                return true;
            }
        }

        if (schoolValue > 0)
        {
            for (var i = 0; i < DiceCount; i++)
                _fixedDice[i] = _dice[i] == schoolValue;
            return true;
        }

        if (bestStraightValues.Count > 0)
        {
            KeepSingleDiceForValues(bestStraightValues);
            return true;
        }

        return false;
    }

    private bool TryKeepCriticalEndgameSchoolDice((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3)
            return false;
        if (CountFreeMainCells() > 8 && !HasHeavySchoolEndgamePressure())
            return false;

        var target = counts
            .Where(item => item.Count > 0
                && CurrentPlayer.GetFreeCell(item.Value - 1) != -1
                && (CurrentPlayer.School < item.Value || HasHeavySchoolEndgamePressure()))
            .OrderByDescending(item => GetCriticalSchoolKeepPriority(item.Value, item.Count))
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();
        if (target.Count == 0)
            return false;
        if (ShouldCriticalSchoolYieldToReadySchoolTriple(counts, target.Value, target.Count))
            return false;

        KeepValues([target.Value]);
        return true;
    }

    private bool ShouldCriticalSchoolYieldToReadySchoolTriple((int Value, int Count)[] counts, int targetValue, int targetCount)
    {
        if (targetCount >= 3)
            return false;

        var readySchoolTriple = counts
            .Where(item => item.Count >= 3 && CurrentPlayer.GetFreeCell(item.Value - 1) != -1)
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();
        if (readySchoolTriple.Count < 3)
            return false;

        var targetRow = targetValue - 1;
        if (targetRow < 0 || targetRow >= 6)
            return false;

        var column = CurrentPlayer.GetFreeCell(targetRow);
        var targetIsImmediateTacticalPrize = CompletesOwnRowPrize(targetRow)
            || ShouldBlockOpponentRowPrize(targetRow)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);
        if (targetIsImmediateTacticalPrize)
            return false;

        return true;
    }

    private bool HasHeavySchoolEndgamePressure()
    {
        var freeMainCells = CountFreeMainCells();
        if (freeMainCells > 18)
            return false;

        var remainingSchoolDemand = Enumerable.Range(1, 6)
            .Where(value => CurrentPlayer.GetFreeCell(value - 1) != -1)
            .Sum(value => (ColumnCount - 1 - CountBusyMainCellsInRow(value - 1)) * value);

        var criticalSchoolRows = Enumerable.Range(1, 6)
            .Where(value => CurrentPlayer.GetFreeCell(value - 1) != -1
                && ColumnCount - 1 - CountBusyMainCellsInRow(value - 1) >= 3)
            .ToArray();
        var criticalSchoolDemand = criticalSchoolRows.Sum(value => (ColumnCount - 1 - CountBusyMainCellsInRow(value - 1)) * value);
        var highSchoolDemand = Enumerable.Range(5, 2)
            .Where(value => CurrentPlayer.GetFreeCell(value - 1) != -1)
            .Sum(value => (ColumnCount - 1 - CountBusyMainCellsInRow(value - 1)) * value);

        return highSchoolDemand >= CurrentPlayer.School + 6
            || criticalSchoolDemand >= CurrentPlayer.School + 10
            || remainingSchoolDemand >= CurrentPlayer.School + 18;
    }

    private bool HasComfortableLeadForSchoolLock()
    {
        var opponent = _players[1 - _currentPlayerIndex];
        var lead = CurrentPlayer.Score - opponent.Score;
        if (lead < 180)
            return false;

        return CountFreeMainCells() <= 16 || HasHeavySchoolEndgamePressure();
    }

    private int GetCriticalSchoolKeepPriority(int value, int count)
    {
        var row = value - 1;
        var remainingCells = ColumnCount - 1 - CountBusyMainCellsInRow(row);
        var priority = count * 24 + value * 14 + remainingCells * value * 3;

        if (value >= 5)
            priority += 90;

        if (CurrentPlayer.School < value)
            priority += 30;

        return priority;
    }

    private int CountFreeMainCells()
    {
        var count = 0;
        for (var row = 0; row < RowCount - 1; row++)
        {
            for (var column = 0; column < ColumnCount - 1; column++)
            {
                if (CurrentPlayer.Table[row, column] == EmptyCell)
                    count++;
            }
        }

        return count;
    }

    private HashSet<int> GetBestEndgameStraightValues()
    {
        var candidates = new[]
        {
            (Row: 10, Values: Enumerable.Range(1, 5).ToHashSet()),
            (Row: 11, Values: Enumerable.Range(2, 5).ToHashSet())
        };

        return candidates
            .Where(candidate => CurrentPlayer.GetFreeCell(candidate.Row) != -1)
            .Select(candidate => new
            {
                candidate.Values,
                Present = candidate.Values.Count(value => _dice.Contains(value))
            })
            .Where(candidate => candidate.Present >= 3)
            .OrderByDescending(candidate => candidate.Present)
            .Select(candidate => candidate.Values)
            .FirstOrDefault() ?? [];
    }

    private int GetBestOpenSchoolValue()
    {
        return Enumerable.Range(1, 6)
            .Where(value => CurrentPlayer.GetFreeCell(value - 1) != -1)
            .OrderByDescending(CountDice)
            .ThenByDescending(value => value)
            .FirstOrDefault();
    }

    private void KeepSingleDiceForValues(HashSet<int> values)
    {
        var kept = new HashSet<int>();
        for (var i = 0; i < DiceCount; i++)
        {
            var value = _dice[i];
            _fixedDice[i] = values.Contains(value) && kept.Add(value);
        }
    }

    // Ищет строку, где текущий ход может закрыть свой приз или помешать ближайшему призу соперника.
    private int GetUrgentComputerTargetRow()
    {
        return Enumerable.Range(0, RowCount - 1)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1 && IsTacticalPrizeMove(row))
            .OrderByDescending(GetUrgentTargetPriority)
            .FirstOrDefault(-1);
    }

    private int GetUrgentTargetPriority(int row)
    {
        var score = CalculateScore(row);
        var priority = GetTacticalPrizeUrgency(row);

        if (score > 0)
            priority += 120 + Math.Min(score, 80);

        priority += row switch
        {
            >= 0 and < 6 => CountDice(row + 1) * 18 + (row + 1),
            6 => PairScore(GetDiceCounts()),
            7 => TwoPairsScore(GetDiceCounts()),
            8 => GetDiceCounts().Values.Max() * 22,
            9 => FullHouseScore(GetDiceCounts()) > 0 ? 70 : GetDiceCounts().Values.Count(count => count >= 2) * 18,
            10 => Enumerable.Range(1, 5).Count(value => _dice.Contains(value)) * 18,
            11 => Enumerable.Range(2, 5).Count(value => _dice.Contains(value)) * 18,
            12 or 13 => GetDiceCounts().Values.Max() * 24,
            14 => Math.Max(0, _dice.Sum() - 16),
            _ => 0
        };

        return priority;
    }

    // Подбирает стратегию фиксации кубиков под конкретную целевую строку таблицы.
    private bool ApplyComputerTargetKeepStrategy(int row, (int Value, int Count)[] counts)
    {
        switch (row)
        {
            case >= 0 and < 6:
                if (CountDice(row + 1) < 2 && HasUsefulPairTarget(counts))
                    return false;

                KeepValues([row + 1]);
                return true;
            case 6:
                return KeepPairs(counts, 1);
            case 7:
                return KeepPairs(counts, 2);
            case 8:
                return KeepMostRepeated(counts);
            case 9:
                return KeepFullHouseCandidate(counts);
            case 10:
                KeepSingleDiceForValues(Enumerable.Range(1, 5).ToHashSet());
                return true;
            case 11:
                KeepSingleDiceForValues(Enumerable.Range(2, 5).ToHashSet());
                return true;
            case 12:
            case 13:
                return KeepMostRepeated(counts);
            case 14:
                return false;
            default:
                return false;
        }
    }

    private bool ShouldPreferStraightForkOverUrgentTarget(int row)
    {
        if (_rollCount != 1 || !HasStraightFork())
            return false;
        if (CurrentPlayer.GetFreeCell(10) == -1 && CurrentPlayer.GetFreeCell(11) == -1)
            return false;
        if (row is < 0 or >= 6)
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        var closesPrizeNow = CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column)
            || ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);
        if (closesPrizeNow)
            return false;

        var schoolPrizeValue = (row + 1) * 5;
        return schoolPrizeValue <= 15;
    }

    private bool ShouldPreferPairOverUrgentStraight(int row, (int Value, int Count)[] counts)
    {
        if (row is not (10 or 11) || HasStraightFork() || !HasUsefulPairTarget(counts))
            return false;
        if (ShouldForceHardTacticalTarget(row))
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        var straightClosesOwnPrize = CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column);

        return !straightClosesOwnPrize;
    }

    private bool ShouldForceHardTacticalTarget(int row)
    {
        if (row is < 9 or > 13)
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        var closesOwnPrize = CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column);
        var blocksOpponentPrize = ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);

        return closesOwnPrize || blocksOpponentPrize;
    }

    private bool TryKeepUsefulPairDice((int Value, int Count)[] counts)
    {
        var pair = counts
            .Where(item => item.Count >= 2 && item.Value >= 2 && HasPairDevelopmentTarget(item.Value))
            .OrderByDescending(item => GetPairKeepPriority(item.Value, item.Count))
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();
        var lockedPairValue = GetLockedUsefulPairValue(counts);
        if (lockedPairValue > 0)
        {
            var lockedPair = counts.FirstOrDefault(item => item.Value == lockedPairValue);
            if (lockedPair.Count >= 2)
            {
                if (pair.Count >= 2
                    && CurrentPlayer.GetFreeCell(lockedPair.Value - 1) != -1
                    && CurrentPlayer.GetFreeCell(pair.Value - 1) == -1
                    && !IsTacticalPrizeMove(8)
                    && !IsTacticalPrizeMove(12)
                    && !IsTacticalPrizeMove(13))
                    pair = lockedPair;

                var lockedPriority = GetPairKeepPriority(lockedPair.Value, lockedPair.Count) + 35;
                var bestPriority = pair.Count >= 2 ? GetPairKeepPriority(pair.Value, pair.Count) : int.MinValue;
                if (lockedPriority >= bestPriority)
                    pair = lockedPair;
            }
        }
        if (pair.Count < 2)
            return false;

        KeepValues([pair.Value]);
        return true;
    }

    private bool TryKeepCriticalSchoolSingleOverDeadPair((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || !HasHeavySchoolEndgamePressure())
            return false;

        var pair = counts
            .Where(item => item.Count >= 2 && item.Value >= 2 && HasPairDevelopmentTarget(item.Value))
            .OrderByDescending(item => GetPairKeepPriority(item.Value, item.Count))
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();
        if (pair.Count < 2)
            return false;
        if (CurrentPlayer.GetFreeCell(pair.Value - 1) != -1)
            return false;

        var schoolSingle = counts
            .Where(item => item.Count > 0 && CurrentPlayer.GetFreeCell(item.Value - 1) != -1)
            .OrderByDescending(item => GetCriticalSchoolKeepPriority(item.Value, item.Count))
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();
        if (schoolSingle.Count == 0)
            return false;

        KeepValues([schoolSingle.Value]);
        return true;
    }

    private bool TryKeepUnderfilledFullHouseCandidate((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || CurrentPlayer.GetFreeCell(9) == -1)
            return false;
        if (CountFreeMainCells() > 14)
            return false;
        var fullRowFill = CountBusyMainCellsInRow(9);
        if (fullRowFill > 2)
            return false;
        if (HasHeavySchoolEndgamePressure())
            return false;
        if (GetComputerSchoolTarget(counts) >= 5)
            return false;
        if (fullRowFill >= 2 && !IsTacticalPrizeMove(9))
            return false;

        var pairValues = counts
            .Where(item => item.Count >= 2)
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .Take(2)
            .Select(item => item.Value)
            .ToArray();
        if (pairValues.Length < 2)
            return false;

        KeepValues(pairValues.ToHashSet());
        return true;
    }

    private bool TryKeepOpenOnesForSchool((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;
        if (CurrentPlayer.GetFreeCell(0) == -1 || CountBusyMainCellsInRow(0) > 2)
            return false;
        if (counts.Any(item => item.Count >= 2 && item.Value >= 2 && HasPairDevelopmentTarget(item.Value)))
            return false;

        var ones = counts.FirstOrDefault(item => item.Value == 1);
        if (ones.Count < 2)
            return false;

        KeepValues([1]);
        return true;
    }

    private bool TryKeepOpenOnesOverClosedSchoolPair((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;
        if (CurrentPlayer.GetFreeCell(0) == -1 || CountBusyMainCellsInRow(0) > 2)
            return false;

        var ones = counts.FirstOrDefault(item => item.Value == 1);
        if (ones.Count < 2)
            return false;

        var competingPair = counts.FirstOrDefault(item =>
            item.Count >= 2
            && item.Value >= 2
            && HasPairDevelopmentTarget(item.Value)
            && CurrentPlayer.GetFreeCell(item.Value - 1) == -1);
        if (competingPair.Count < 2)
            return false;

        KeepValues([1]);
        return true;
    }

    private bool TryKeepStrongOnesBeforeUrgentTarget((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;
        if (CurrentPlayer.GetFreeCell(0) == -1)
            return false;

        var ones = counts.FirstOrDefault(item => item.Value == 1);
        if (ones.Count < 3)
            return false;

        var urgentTargetRow = GetUrgentComputerTargetRow();
        if (urgentTargetRow >= 0 && UrgentTargetMustOverrideOnes(urgentTargetRow))
            return false;

        KeepValues([1]);
        return true;
    }

    private bool TryKeepOpenOnesBeforeUrgentTarget((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;
        if (CurrentPlayer.GetFreeCell(0) == -1 || CountBusyMainCellsInRow(0) > 2)
            return false;
        if (counts.Any(item => item.Count >= 2 && item.Value >= 2 && HasPairDevelopmentTarget(item.Value)))
            return false;

        var ones = counts.FirstOrDefault(item => item.Value == 1);
        if (ones.Count < 2)
            return false;

        var urgentTargetRow = GetUrgentComputerTargetRow();
        if (urgentTargetRow < 0)
            return false;
        if (!ShouldPreferOpenOnesOverUrgentTarget(urgentTargetRow))
            return false;

        KeepValues([1]);
        return true;
    }

    private bool ShouldPreferOpenOnesOverUrgentTarget(int row)
    {
        if (row is < 0 or >= 6)
            return false;
        if (ShouldBlockOpponentRowPrize(row))
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        if (column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column))
            return false;

        var schoolPrizeValue = (row + 1) * 5;
        return CountDice(row + 1) <= 1 && schoolPrizeValue <= 15;
    }

    private bool UrgentTargetMustOverrideOnes(int row)
    {
        if (row is < 0 or >= RowCount)
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        return CompletesOwnRowPrize(row)
            || ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);
    }

    private bool HasUsefulPairTarget((int Value, int Count)[] counts)
    {
        return counts.Any(item => item.Count >= 2 && item.Value >= 2 && HasPairDevelopmentTarget(item.Value));
    }

    private bool TryKeepOpenSchoolPairBeforeUrgentTarget((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;

        var schoolPair = counts
            .Where(item => item.Count >= 2
                && CurrentPlayer.GetFreeCell(item.Value - 1) != -1)
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => GetOpenSchoolPairPriority(item.Value))
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();
        if (schoolPair.Count < 2)
            return false;

        if (schoolPair.Value == 1 && CountBusyMainCellsInRow(0) <= 1)
        {
            KeepValues([1]);
            return true;
        }

        var criticalSchoolSingle = GetCriticalSchoolSingleCandidate(counts, schoolPair.Value);
        if (ShouldPreferCriticalSchoolSingleOverSchoolPair(schoolPair.Value, schoolPair.Count, criticalSchoolSingle))
        {
            KeepValues([criticalSchoolSingle.Value]);
            return true;
        }

        var urgentTargetRow = GetUrgentComputerTargetRow();
        if (urgentTargetRow >= 0 && UrgentTargetMustOverrideSchoolPair(urgentTargetRow, schoolPair.Value))
            return false;

        KeepValues([schoolPair.Value]);
        return true;
    }

    private int GetOpenSchoolPairPriority(int value)
    {
        var row = value - 1;
        var rowFill = CountBusyMainCellsInRow(row);
        return value * 20 - rowFill * 12 + GetSchoolPrizeRaceBonus(row);
    }

    private (int Value, int Count) GetCriticalSchoolSingleCandidate((int Value, int Count)[] counts, int excludedValue)
    {
        return counts
            .Where(item => item.Count > 0
                && item.Value != excludedValue
                && CurrentPlayer.GetFreeCell(item.Value - 1) != -1)
            .OrderByDescending(item => GetCriticalSchoolKeepPriority(item.Value, item.Count))
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();
    }

    private bool ShouldPreferCriticalSchoolSingleOverSchoolPair(int pairValue, int pairCount, (int Value, int Count) schoolSingle)
    {
        if (pairValue < 2 || pairCount < 2 || schoolSingle.Count == 0)
            return false;
        if (!HasHeavySchoolEndgamePressure() && CountFreeMainCells() > 10)
            return false;

        var pairRow = pairValue - 1;
        var pairColumn = CurrentPlayer.GetFreeCell(pairRow);
        var pairIsImmediateTactical = CompletesOwnRowPrize(pairRow)
            || ShouldBlockOpponentRowPrize(pairRow)
            || pairColumn >= 0 && pairColumn < ColumnCount - 1 && CompletesOwnColumnPrize(pairColumn)
            || pairColumn >= 0 && pairColumn < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(pairColumn);
        if (pairIsImmediateTactical)
            return false;

        var singlePriority = GetCriticalSchoolKeepPriority(schoolSingle.Value, schoolSingle.Count);
        var pairPriority = GetOpenSchoolPairPriority(pairValue) + pairCount * 24;
        var singleNeeds = ColumnCount - 1 - CountBusyMainCellsInRow(schoolSingle.Value - 1);
        var pairNeeds = ColumnCount - 1 - CountBusyMainCellsInRow(pairRow);

        return singlePriority >= pairPriority + 40
            || schoolSingle.Value >= 5 && singlePriority >= pairPriority + 20
            || singleNeeds >= 3 && pairNeeds <= 2 && singlePriority >= pairPriority;
    }

    private bool TryKeepReadySchoolTripleBeforeUrgentTarget((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;

        var schoolTriple = counts
            .Where(item => item.Count >= 3
                && CurrentPlayer.GetFreeCell(item.Value - 1) != -1)
            .OrderByDescending(item => item.Value)
            .FirstOrDefault();
        if (schoolTriple.Count < 3)
            return false;

        var urgentTargetRow = GetUrgentComputerTargetRow();
        if (urgentTargetRow >= 0 && UrgentTargetMustOverrideSchoolTriple(urgentTargetRow, schoolTriple.Value))
            return false;

        KeepValues([schoolTriple.Value]);
        return true;
    }

    private bool TryKeepOpeningTripleDevelopment((int Value, int Count)[] counts)
    {
        if (_rollCount != 1 || HasStraightFork())
            return false;

        var openingTriple = counts
            .Where(item => item.Count >= 3 && HasOpeningTripleDevelopmentTarget(item.Value))
            .OrderByDescending(item => GetOpeningTripleKeepPriority(item.Value))
            .FirstOrDefault();
        if (openingTriple.Count < 3)
            return false;

        var urgentTargetRow = GetUrgentComputerTargetRow();
        if (urgentTargetRow >= 0 && UrgentTargetMustOverrideSchoolTriple(urgentTargetRow, openingTriple.Value))
            return false;

        KeepValues([openingTriple.Value]);
        return true;
    }

    private bool HasOpeningTripleDevelopmentTarget(int value)
    {
        var schoolRow = value - 1;
        return CurrentPlayer.GetFreeCell(schoolRow) != -1
            || CurrentPlayer.GetFreeCell(8) != -1
            || CurrentPlayer.GetFreeCell(12) != -1
            || CurrentPlayer.GetFreeCell(13) != -1
            || CurrentPlayer.GetFreeCell(9) != -1;
    }

    private int GetOpeningTripleKeepPriority(int value)
    {
        var priority = value * 30;

        if (CurrentPlayer.GetFreeCell(value - 1) != -1)
            priority += 80;
        if (CurrentPlayer.GetFreeCell(8) != -1)
            priority += 60;
        if (CurrentPlayer.GetFreeCell(12) != -1)
            priority += 40;
        if (CurrentPlayer.GetFreeCell(13) != -1)
            priority += 30;

        priority += value switch
        {
            6 => 40,
            5 => 30,
            4 => 16,
            _ => 0
        };

        return priority;
    }

    private bool UrgentTargetMustOverrideSchoolTriple(int row, int schoolValue)
    {
        if (row == schoolValue - 1)
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        var ownPrizeNow = CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column);
        var blocksOpponentColumnNow = column >= 0
            && column < ColumnCount - 1
            && ShouldBlockOpponentColumnPrize(column);

        if (row < 6
            && CountDice(schoolValue) >= 3
            && CountDice(row + 1) < 3
            && !ownPrizeNow
            && !blocksOpponentColumnNow)
            return false;

        return CompletesOwnRowPrize(row)
            || ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);
    }

    private bool UrgentTargetMustOverrideSchoolPair(int row, int schoolValue)
    {
        if (row == schoolValue - 1)
            return false;

        var column = CurrentPlayer.GetFreeCell(row);
        var ownPrizeNow = CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column);
        var blocksOpponentColumnNow = column >= 0
            && column < ColumnCount - 1
            && ShouldBlockOpponentColumnPrize(column);

        if (row < 6
            && CountDice(schoolValue) >= 2
            && CountDice(row + 1) < 3
            && !ownPrizeNow
            && !blocksOpponentColumnNow)
            return false;

        if (row == 8
            && CountDice(schoolValue) >= 2
            && CurrentPlayer.GetFreeCell(schoolValue - 1) != -1
            && !ownPrizeNow
            && !ShouldBlockOpponentRowPrize(row)
            && !blocksOpponentColumnNow)
            return false;

        return CompletesOwnRowPrize(row)
            || ShouldBlockOpponentRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column)
            || column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column);
    }

    private bool HasPairDevelopmentTarget(int value)
    {
        return CurrentPlayer.GetFreeCell(value - 1) != -1
            || CurrentPlayer.GetFreeCell(6) != -1
            || CurrentPlayer.GetFreeCell(7) != -1
            || CurrentPlayer.GetFreeCell(8) != -1
            || CurrentPlayer.GetFreeCell(9) != -1
            || CurrentPlayer.GetFreeCell(12) != -1
            || CurrentPlayer.GetFreeCell(13) != -1;
    }

    private int GetLockedUsefulPairValue((int Value, int Count)[] counts)
    {
        if (!_fixedDice.Any(isFixed => isFixed))
            return 0;

        var lockedValues = Enumerable.Range(0, DiceCount)
            .Where(index => _fixedDice[index])
            .Select(index => _dice[index])
            .Distinct()
            .ToArray();
        if (lockedValues.Length != 1)
            return 0;

        var lockedValue = lockedValues[0];
        if (lockedValue < 2 || !HasPairDevelopmentTarget(lockedValue))
            return 0;

        return counts.Any(item => item.Value == lockedValue && item.Count >= 2)
            ? lockedValue
            : 0;
    }

    private bool TryKeepHighSingleInsteadOfWeakLowPair((int Value, int Count)[] counts)
    {
        if (_rollCount >= 3 || HasStraightFork())
            return false;
        if (counts.Any(item => item.Count >= 2 && item.Value >= 2 && HasPairDevelopmentTarget(item.Value)))
            return false;
        if (GetComputerSchoolTarget(counts) > 0)
            return false;

        var repeatedLowValue = counts
            .Where(item => item.Count >= 2)
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .Select(item => item.Value)
            .FirstOrDefault();
        if (repeatedLowValue is 0 or > 2)
            return false;
        if (CurrentPlayer.GetFreeCell(repeatedLowValue - 1) != -1)
            return false;

        var highValue = counts
            .Where(item => item.Value >= 5 && item.Count == 1 && HasPairDevelopmentTarget(item.Value))
            .OrderByDescending(item => item.Value)
            .Select(item => item.Value)
            .FirstOrDefault();
        if (highValue == 0)
            return false;

        KeepValues([highValue]);
        return true;
    }

    private bool KeepPairs((int Value, int Count)[] counts, int pairCount)
    {
        var pairValues = counts
            .Where(item => item.Count >= 2)
            .OrderByDescending(item => GetPairKeepPriority(item.Value, item.Count))
            .ThenByDescending(item => item.Value)
            .Take(pairCount)
            .Select(item => item.Value)
            .ToHashSet();

        if (pairValues.Count == 0)
            pairValues.Add(counts.OrderByDescending(item => item.Value).First().Value);

        KeepValues(pairValues);
        return true;
    }

    private int GetPairKeepPriority(int value, int count)
    {
        var schoolRow = value - 1;
        var schoolIsOpen = CurrentPlayer.GetFreeCell(schoolRow) != -1;
        var schoolFill = schoolIsOpen ? CountBusyMainCellsInRow(schoolRow) : ColumnCount - 1;
        var priority = count * 100 - schoolFill * 18 + value;

        if (schoolIsOpen)
            priority += 30;
        if (IsSchoolPrizeRaceTarget(schoolRow))
            priority += 80;

        priority += GetSchoolPairTacticalKeepBonus(schoolRow);

        return priority;
    }

    private int GetSchoolPairTacticalKeepBonus(int row)
    {
        if (row is < 0 or >= 6 || CurrentPlayer.GetFreeCell(row) == -1)
            return 0;

        var column = CurrentPlayer.GetFreeCell(row);
        var bonus = 0;

        if (ShouldBlockOpponentRowPrize(row))
            bonus += 120;
        if (column >= 0 && column < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(column))
            bonus += 120;
        if (CompletesOwnRowPrize(row))
            bonus += 80;
        if (column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column))
            bonus += 80;

        return bonus;
    }

    private bool KeepMostRepeated((int Value, int Count)[] counts)
    {
        KeepValues([counts[0].Value]);
        return true;
    }

    private bool KeepFullHouseCandidate((int Value, int Count)[] counts)
    {
        var keepValues = counts
            .Where(item => item.Count >= 2)
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .Take(2)
            .Select(item => item.Value)
            .ToHashSet();

        if (keepValues.Count == 0)
            keepValues.Add(counts[0].Value);

        KeepValues(keepValues);
        return true;
    }

    private void KeepValues(HashSet<int> values)
    {
        for (var i = 0; i < DiceCount; i++)
            _fixedDice[i] = values.Contains(_dice[i]);
    }

    // Проверяет, есть ли смысл ловить стрит.
    // Если уже есть пара/тройка, обычный стрит не навязывается, кроме симметричной основы 2-3-4-5.
    private bool HasOpenStraightTarget(int smallStraight, int bigStraight, int straightTarget, IEnumerable<(int Value, int Count)> counts)
    {
        if (HasStraightFork())
            return CurrentPlayer.GetFreeCell(10) != -1 || CurrentPlayer.GetFreeCell(11) != -1;

        var tacticalSmallStraight = CurrentPlayer.GetFreeCell(10) != -1
            && smallStraight >= straightTarget
            && IsTacticalPrizeMove(10);
        var tacticalBigStraight = CurrentPlayer.GetFreeCell(11) != -1
            && bigStraight >= straightTarget
            && IsTacticalPrizeMove(11);

        if (counts.Any(item => item.Count >= 2)
            && !tacticalSmallStraight
            && !tacticalBigStraight)
            return false;

        return tacticalSmallStraight || tacticalBigStraight;
    }

    private HashSet<int> GetStraightKeepValues(int smallStraight, int bigStraight)
    {
        if (HasStraightFork())
            return Enumerable.Range(2, 4).ToHashSet();

        if (CurrentPlayer.GetFreeCell(10) == -1)
            return Enumerable.Range(2, 5).ToHashSet();

        if (CurrentPlayer.GetFreeCell(11) == -1)
            return Enumerable.Range(1, 5).ToHashSet();

        return smallStraight >= bigStraight
            ? Enumerable.Range(1, 5).ToHashSet()
            : Enumerable.Range(2, 5).ToHashSet();
    }

    private bool HasStraightFork()
    {
        return Enumerable.Range(2, 4).All(value => _dice.Contains(value));
    }

    private void KeepStraightForkDice()
    {
        var needed = Enumerable.Range(2, 4).ToHashSet();
        var alreadyKept = new HashSet<int>();

        for (var i = 0; i < DiceCount; i++)
        {
            var die = _dice[i];
            var keep = needed.Contains(die) && alreadyKept.Add(die);
            _fixedDice[i] = keep;
        }
    }

    // Решает, стоит ли продолжать ловить конкретную строку школы.
    // При перекосе в школу компьютер переключается на комбинации, кроме очень сильных школьных номиналов.
    private bool ShouldChaseSchoolValue(int value, int count, bool preferCombinations)
    {
        if (IsSchoolPrizeRaceTarget(value - 1))
            return true;

        if (count >= 3)
            return CurrentPlayer.School <= 0 || !preferCombinations || value >= 5;

        if (count < 2)
            return false;

        if (preferCombinations)
            return false;

        if (CurrentPlayer.School >= 0)
            return value >= 5 && CurrentPlayer.Difficulty == ComputerDifficulty.Careful;

        return true;
    }

    private int GetSchoolPrizeRaceBonus(int row)
    {
        if (!IsSchoolPrizeRaceTarget(row))
            return 0;

        var rowFill = CountBusyMainCellsInRow(row);
        return 70 + rowFill * 18;
    }

    private bool IsSchoolPrizeRaceTarget(int row)
    {
        if (row is < 0 or >= 6 || CurrentPlayer.GetFreeCell(row) == -1)
            return false;
        if (CurrentPlayer.Table[row, ColumnCount - 1] != EmptyCell)
            return false;

        var rowFill = CountBusyMainCellsInRow(row);
        var requiredFill = row >= 4 ? 3 : 4;
        return rowFill >= requiredFill && OpponentPrizeIsOpen(row, ColumnCount - 1)
            || OpponentCanClaimRowPrizeSoon(row);
    }

    private readonly record struct ComputerMove(int Row, int Score, int AdjustedScore);
}
