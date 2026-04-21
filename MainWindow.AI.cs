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

        if (CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive && IsTacticalPrizeMove(best.Row))
            return true;

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

        if (_rollCount == 1)
            return best.Row == 13
                || best.Row == 12 && best.Score > 0
                || best.Row == 7 && (HasPremiumTwoPairs(GetDiceCounts()) || IsTacticalPrizeMove(7))
                || best.Row == 8 && best.Score > 0 && IsStrongOpeningTriple(GetDiceCounts()) && !HasStraightFork()
                || best.Row == 9 && best.Score > 0
                || best.Row is 10 or 11 && best.Score > 0;

        if (best.Row is 10 or 11 && best.Score > 0)
            return true;

        if (best.Row == 9 && _rollCount >= 2 && best.Score >= 24)
            return true;

        if (best.Row < 6 && CountDice(best.Row + 1) == 3)
            return true;

        if (best.Row == 13 || best.Row == 12 && best.Score > 0)
            return true;

        return best.Score >= threshold;
    }

    // Главный выбор строки для записи результата.
    // Сначала идут жесткие тактические правила, которые должны перебить обычные веса:
    // сильная готовая комбинация, дешевый минус в школе, пара вместо слабых двух пар, школа после третьего броска.
    // Если жесткого правила нет, все свободные строки оцениваются общей функцией приоритета.
    private ComputerMove GetBestComputerMove()
    {
        var forcedPremiumRow = GetForcedPremiumMove();
        if (forcedPremiumRow >= 0)
            return new ComputerMove(forcedPremiumRow, CalculateScore(forcedPremiumRow), int.MaxValue);

        var forcedTacticalPrizeRow = GetForcedTacticalPrizeMove();
        if (forcedTacticalPrizeRow >= 0)
            return new ComputerMove(forcedTacticalPrizeRow, CalculateScore(forcedTacticalPrizeRow), int.MaxValue);

        var forcedLowCostSchoolRow = GetForcedLowCostSchoolMove();
        if (forcedLowCostSchoolRow >= 0)
            return new ComputerMove(forcedLowCostSchoolRow, CalculateScore(forcedLowCostSchoolRow), int.MaxValue);

        var forcedSchoolRow = GetForcedSchoolMove();
        if (forcedSchoolRow >= 0)
            return new ComputerMove(forcedSchoolRow, CalculateScore(forcedSchoolRow), int.MaxValue);

        var forcedTripleRow = GetForcedTripleMove();
        if (forcedTripleRow >= 0)
            return new ComputerMove(forcedTripleRow, CalculateScore(forcedTripleRow), int.MaxValue);

        var forcedWeakResultCrossOutRow = GetForcedWeakResultCrossOutMove();
        if (forcedWeakResultCrossOutRow >= 0)
            return new ComputerMove(forcedWeakResultCrossOutRow, CalculateScore(forcedWeakResultCrossOutRow), int.MaxValue);

        var forcedWeakPairRow = GetForcedWeakPairMove();
        if (forcedWeakPairRow >= 0)
            return new ComputerMove(forcedWeakPairRow, CalculateScore(forcedWeakPairRow), int.MaxValue);

        var bestRow = -1;
        var bestAdjusted = int.MinValue;

        // Общая оценка всех доступных строк: реальные очки хода плюс бонусы/штрафы за состояние таблицы.
        // Здесь же отсекаются запрещенные школьные минусы и слишком дорогие списания школы.
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
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

        var aggressivePrizeRaceRow = GetAggressiveOpeningPrizeRaceMove(counts);
        if (aggressivePrizeRaceRow >= 0)
            return aggressivePrizeRaceRow;

        if (CurrentPlayer.GetFreeCell(7) != -1 && TwoPairsScore(counts) > 0 && IsTacticalPrizeMove(7))
            return 7;

        if (CurrentPlayer.GetFreeCell(7) != -1 && TwoPairsScore(counts) >= 18)
            return 7;

        if (CurrentPlayer.GetFreeCell(6) != -1 && PairScore(counts) >= 10 && ShouldTakeOpeningPair())
            return 6;

        if (CurrentPlayer.GetFreeCell(8) != -1 && counts.Any(pair => pair.Value >= 3))
            return 8;

        if (ShouldTakeOpeningSum(counts))
            return 14;

        if (CurrentPlayer.GetFreeCell(7) != -1 && HasPremiumTwoPairs(counts))
            return 7;

        if (CurrentPlayer.GetFreeCell(14) != -1 && IsGoodOpeningSum() && OpponentCanClaimRowPrizeSoon(14))
            return 14;

        if (CurrentPlayer.GetFreeCell(14) != -1 && IsGoodOpeningSum())
            return 14;

        return -1;
    }

    // Проверка тактической ценности хода: свой приз или блокировка приза соперника
    // учитываются и по строке таблицы, и по столбцу таблицы.
    private bool IsTacticalPrizeMove(int row)
    {
        var column = CurrentPlayer.GetFreeCell(row);
        return CompletesOwnRowPrize(row)
            || column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column)
            || OpponentCanClaimRowPrizeSoon(row)
            || column >= 0 && column < ColumnCount - 1 && OpponentCanClaimColumnPrizeSoon(column);
    }

    private int GetForcedTacticalPrizeMove()
    {
        return Enumerable.Range(0, RowCount - 1)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1 && IsTacticalPrizeMove(row))
            .Select(row => (Row: row, Score: CalculateScore(row)))
            .Where(item => IsPlayableTacticalPrizeScore(item.Row, item.Score))
            .OrderByDescending(item => GetTacticalPrizeUrgency(item.Row))
            .ThenByDescending(item => item.Score)
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    private bool IsPlayableTacticalPrizeScore(int row, int score)
    {
        if (row < 6)
            return CountDice(row + 1) > 0 && (score >= 0 || -score <= CurrentPlayer.School);

        return score > 0;
    }

    private int GetTacticalPrizeUrgency(int row)
    {
        var column = CurrentPlayer.GetFreeCell(row);
        var urgency = 0;
        if (CompletesOwnRowPrize(row))
            urgency += OpponentCanClaimRowPrizeSoon(row) ? 140 : 90;
        if (OpponentCanClaimRowPrizeSoon(row))
            urgency += 120;
        if (column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column))
            urgency += OpponentCanClaimColumnPrizeSoon(column) ? 100 : 60;
        if (column >= 0 && column < ColumnCount - 1 && OpponentCanClaimColumnPrizeSoon(column))
            urgency += 80;

        return urgency;
    }

    // После третьего броска ИИ может принудительно записать школу,
    // если нет более ценной готовой комбинации и ровно три кости нужного номинала уже собраны.
    private int GetForcedSchoolMove()
    {
        if (_rollCount < 3)
            return -1;

        return GetDiceCounts()
            .Where(pair => pair.Value >= 3 && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1)
            .OrderByDescending(pair => pair.Key)
            .Select(pair => pair.Key - 1)
            .FirstOrDefault(-1);
    }

    // Если последним броском получилась только слабая комбинация,
    // маленький допустимый минус в школе может быть лучше, чем тратить важную строку комбинаций.
    private int GetForcedLowCostSchoolMove()
    {
        if (_rollCount < 3 || CurrentPlayer.School < 6)
            return -1;

        var bestWeakCombination = GetBestWeakCombinationMove();
        if (bestWeakCombination.Row == -1 || bestWeakCombination.Score > 10 || IsTacticalPrizeMove(bestWeakCombination.Row))
            return -1;

        return Enumerable.Range(0, 3)
            .Where(row => CurrentPlayer.GetFreeCell(row) != -1)
            .Select(row => (Row: row, Score: CalculateScore(row), DiceCount: CountDice(row + 1)))
            .Where(item => item.DiceCount > 0 && item.Score < 0 && ShouldAllowNegativeSchool(item.Row, item.Score))
            .OrderByDescending(item => item.Score)
            .ThenBy(item => item.Row)
            .Select(item => item.Row)
            .FirstOrDefault(-1);
    }

    // Для слабых двух пар меньше 18 очков ИИ сначала пробует школу,
    // а если школа недоступна, выбирает пару, чтобы сохранить строку "две пары" для более сильного хода.
    private int GetForcedWeakPairMove()
    {
        if (_rollCount < 3 || CurrentPlayer.GetFreeCell(6) == -1)
            return -1;

        if (HasBetterFinishedMoveThanWeakTwoPairs())
            return -1;

        var twoPairsScore = CalculateScore(7);
        if (twoPairsScore is <= 0 or >= 18 || IsTacticalPrizeMove(7))
            return -1;

        return CalculateScore(6) > 0 ? 6 : -1;
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

        if (score >= 22)
            return true;

        var fullRowIsAlreadyUsed = CountBusyMainCellsInRow(9) >= 2;
        if (CurrentPlayer.School <= 6 || fullRowIsAlreadyUsed)
            return false;

        return false;
    }

    private bool ShouldTakeOpeningPair()
    {
        return _rollCount == 1
            && CountBusyMainCellsInRow(6) == 0
            && ShouldPreferCombinationsOverSchool();
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

    private int GetForcedWeakResultCrossOutMove()
    {
        if (_rollCount < 3 || !HasManyFreeCombinationCells())
            return -1;

        var pairScore = CurrentPlayer.GetFreeCell(6) != -1 && !IsTacticalPrizeMove(6)
            ? CalculateScore(6)
            : 0;
        var twoPairsScore = CurrentPlayer.GetFreeCell(7) != -1 && !IsTacticalPrizeMove(7)
            ? CalculateScore(7)
            : 0;

        var hasOnlyWeakPairResult = pairScore is > 0 and <= 8;
        var hasOnlyWeakTwoPairsResult = twoPairsScore is > 0 and < 18;
        if (!hasOnlyWeakPairResult && !hasOnlyWeakTwoPairsResult)
            return -1;

        return GetSacrificialCrossOutRow();
    }

    private bool HasManyFreeCombinationCells()
    {
        var freeCells = 0;
        for (var row = 6; row < RowCount - 1; row++)
        {
            for (var column = 0; column < ColumnCount - 1; column++)
            {
                if (CurrentPlayer.Table[row, column] == EmptyCell)
                    freeCells++;
            }
        }

        return freeCells >= 20;
    }

    private int GetSacrificialCrossOutRow()
    {
        return new[] { 13, 11, 10 }
            .Select((row, index) => new
            {
                Row = row,
                Order = index,
                Column = CurrentPlayer.GetFreeCell(row),
                RowCrosses = CountCrossesInMainRow(row)
            })
            .Where(item => item.Column >= 0 && item.Column < ColumnCount - 2 && CalculateScore(item.Row) == 0)
            .Select(item => new
            {
                item.Row,
                Priority = CountCrossesInMainColumn(item.Column) * 40
                    - item.RowCrosses * 24
                    - item.Order
            })
            .OrderByDescending(item => item.Priority)
            .Select(item => item.Row)
            .FirstOrDefault(-1);
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

        return !hasStrongReadyCombination;
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

        if (TryKeepEndgameTargetDice())
            return;

        // Если строка или столбец почти закрывают приз, сначала пытаемся удерживать кости под этот срочный ряд.
        var urgentTargetRow = GetUrgentComputerTargetRow();
        if (urgentTargetRow >= 0 && ApplyComputerTargetKeepStrategy(urgentTargetRow, counts))
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
            for (var i = 0; i < DiceCount; i++)
                _fixedDice[i] = straightValues.Contains(_dice[i]);
            return;
        }

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
                _fixedDice[i] = _dice[i] == pairValue || (_rollCount >= 2 && _dice[i] >= 5);
            return;
        }

        var keepValue = counts[0].Value;
        if (CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive && counts[0].Count < 3)
            keepValue = counts.OrderByDescending(item => item.Value).First().Value;

        for (var i = 0; i < DiceCount; i++)
            _fixedDice[i] = _dice[i] == keepValue;
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
        for (var row = 0; row < RowCount - 1; row++)
        {
            var column = CurrentPlayer.GetFreeCell(row);
            if (column == -1)
                continue;

            if (CompletesOwnRowPrize(row))
                return row;

            if (column < ColumnCount - 1 && CompletesOwnColumnPrize(column))
                return row;

            if (OpponentCanClaimRowPrizeSoon(row))
                return row;

            if (column < ColumnCount - 1 && OpponentCanClaimColumnPrizeSoon(column))
                return row;
        }

        return -1;
    }

    // Подбирает стратегию фиксации кубиков под конкретную целевую строку таблицы.
    private bool ApplyComputerTargetKeepStrategy(int row, (int Value, int Count)[] counts)
    {
        switch (row)
        {
            case >= 0 and < 6:
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
                KeepValues(Enumerable.Range(1, 5).ToHashSet());
                return true;
            case 11:
                KeepValues(Enumerable.Range(2, 5).ToHashSet());
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

    private bool KeepPairs((int Value, int Count)[] counts, int pairCount)
    {
        var pairValues = counts
            .Where(item => item.Count >= 2)
            .OrderByDescending(item => item.Value)
            .Take(pairCount)
            .Select(item => item.Value)
            .ToHashSet();

        if (pairValues.Count == 0)
            pairValues.Add(counts.OrderByDescending(item => item.Value).First().Value);

        KeepValues(pairValues);
        return true;
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

        if (counts.Any(item => item.Count >= 2))
            return false;

        return CurrentPlayer.GetFreeCell(10) != -1 && smallStraight >= straightTarget && IsTacticalPrizeMove(10)
            || CurrentPlayer.GetFreeCell(11) != -1 && bigStraight >= straightTarget && IsTacticalPrizeMove(11);
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
        return rowFill >= 3 && OpponentPrizeIsOpen(row, ColumnCount - 1)
            || OpponentCanClaimRowPrizeSoon(row);
    }

    private readonly record struct ComputerMove(int Row, int Score, int AdjustedScore);
}
