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

        if (_rollCount == 1)
            return best.Row == 13
                || best.Row == 12 && best.Score > 0
                || best.Row == 7 && HasPremiumTwoPairs(GetDiceCounts())
                || best.Row == 8 && best.Score > 0 && !HasStraightFork()
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

        var forcedLowCostSchoolRow = GetForcedLowCostSchoolMove();
        if (forcedLowCostSchoolRow >= 0)
            return new ComputerMove(forcedLowCostSchoolRow, CalculateScore(forcedLowCostSchoolRow), int.MaxValue);

        var forcedWeakPairRow = GetForcedWeakPairMove();
        if (forcedWeakPairRow >= 0)
            return new ComputerMove(forcedWeakPairRow, CalculateScore(forcedWeakPairRow), int.MaxValue);

        var forcedSchoolRow = GetForcedSchoolMove();
        if (forcedSchoolRow >= 0)
            return new ComputerMove(forcedSchoolRow, CalculateScore(forcedSchoolRow), int.MaxValue);

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

    // Мягкая оценка хода для ИИ.
    // Она не запрещает ход напрямую, а добавляет/снимает очки приоритета:
    // школа, баланс комбинаций, бонусы раздачи, призы по строкам/столбцам и заполнение таблицы.
    private int GetComputerMovePriority(int row, int score)
    {
        var priority = score + GetDifficultyRowBonus(row, score);
        var counts = GetDiceCounts();
        var column = CurrentPlayer.GetFreeCell(row);

        if (row < 6)
        {
            var value = row + 1;
            var count = counts[value];
            if (score >= 0)
            {
                priority += 22 + value * 3;
                if (count >= 3)
                    priority += 38 + value * 2;
                if (CurrentPlayer.Table[row, 0] == EmptyCell)
                    priority += 12;
            }
        }
        else if (row < RowCount - 1)
        {
            priority += GetCombinationBalanceBonus(row, score);
        }

        if (_rollCount == 1)
        {
            if (row == 7 && score > 0)
                priority += HasPremiumTwoPairs(counts) ? 54 : score >= 16 ? 12 : -10;
            if (row == 8 && score > 0)
                priority += counts.Any(pair => pair.Value == 3 && pair.Key >= 5) ? 62 : 40;
            if (row == 9 && score > 0)
                priority += 96;
            if (row is 10 or 11 && score > 0)
                priority += 82;
            if (row == 12 && score > 0)
                priority += 120;
            if (row == 13 && score > 0)
                priority += 160;
        }

        if (row == 14)
            priority += _rollCount == 1 && IsGoodOpeningSum() ? 18 : _rollCount < 3 ? -36 : -16;
        if (row == 14 && score < 21 && HasSacrificialCrossOutAvailable())
            priority -= 34;

        if (CompletesOwnRowPrize(row))
            priority += OpponentCanClaimRowPrizeSoon(row) ? 72 : OpponentPrizeIsOpen(row, ColumnCount - 1) ? 34 : 16;
        if (column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column))
            priority += OpponentCanClaimColumnPrizeSoon(column) ? 64 : OpponentPrizeIsOpen(RowCount - 1, column) ? 30 : 14;
        if (column >= 0 && column < ColumnCount - 1)
            priority += GetTableFillBalanceBonus(row, column, score);

        return priority;
    }

    // Подталкивает ИИ к комбинациям, если школа уже слишком заполнена.
    // Это защита от перекоса, когда компьютер играет почти только школу и оставляет строки комбинаций пустыми.
    private int GetCombinationBalanceBonus(int row, int score)
    {
        var bonus = 0;
        if (ShouldPreferCombinationsOverSchool())
            bonus += score > 0 ? 34 : 18;

        if (row == 12 && CountBusyMainCellsInRow(row) == 0)
            bonus += score > 0 ? 42 : 16;

        if (row is 10 or 11 && CountBusyMainCellsInRow(row) == 0)
            bonus += score > 0 ? 18 : 8;

        return bonus;
    }

    // Слабый балансировщик таблицы: выгоднее закрывать менее заполненные строки и столбцы,
    // но без жесткого запрета на другие ходы.
    private int GetTableFillBalanceBonus(int row, int column, int score)
    {
        var bonus = 0;
        var rowBusy = CountBusyMainCellsInRow(row);
        var leastBusyRow = Enumerable.Range(0, RowCount - 1).Min(CountBusyMainCellsInRow);
        if (rowBusy <= leastBusyRow + 1)
            bonus += score > 0 ? 14 : 8;
        else if (rowBusy >= 4 && !CompletesOwnRowPrize(row))
            bonus -= 14;

        var columnBusy = CountBusyMainCellsInColumn(column);
        var leastBusyColumn = Enumerable.Range(0, ColumnCount - 1).Min(CountBusyMainCellsInColumn);
        if (columnBusy <= leastBusyColumn + 2)
            bonus += score > 0 ? 18 : 10;
        else if (columnBusy >= 12 && !CompletesOwnColumnPrize(column))
            bonus -= 12;

        return bonus;
    }

    private int GetDifficultyRowBonus(int row, int score)
    {
        return CurrentPlayer.Difficulty switch
        {
            ComputerDifficulty.Careful when row < 6 && score >= 0 => 12,
            ComputerDifficulty.Careful when row == 14 => 4,
            ComputerDifficulty.Aggressive when row is >= 10 and <= 13 => 12,
            ComputerDifficulty.Aggressive when row >= 6 && score > 0 => 5,
            _ => 0
        };
    }

    // Штраф за зачеркивание комбинации вместо записи очков.
    // Редкие, но маловероятные строки можно вычеркивать раньше; каре, тройку и последний игровой столбец бережем.
    private int GetComputerCrossOutPenalty(int row)
    {
        var penalty = row switch
        {
            13 => 0,  // абак: редкая комбинация, ее можно вычеркивать первой
            11 => 3,  // большой стрит
            10 => 4,  // малый стрит
            9 => 10,  // фул
            7 => 13,  // две пары
            6 => 15,  // пара
            12 => 28, // каре лучше беречь
            8 => 30,  // тройку тоже не вычеркиваем рано
            _ => 18
        };

        if (CompletesOwnRowPrize(row))
            penalty -= 10;

        penalty += CountCrossesInMainRow(row) * 24;
        var column = CurrentPlayer.GetFreeCell(row);
        if (column >= 0 && column < ColumnCount - 1)
            penalty += CountCrossesInMainColumn(column) * 6;
        if (column == ColumnCount - 2)
            penalty += 80;

        return Math.Max(0, penalty);
    }

    private bool HasSacrificialCrossOutAvailable()
    {
        return IsRowOpenForCrossOut(13) || IsRowOpenForCrossOut(11) || IsRowOpenForCrossOut(10);
    }

    private bool IsRowOpenForCrossOut(int row)
    {
        return CurrentPlayer.GetFreeCell(row) != -1 && CalculateScore(row) == 0;
    }

    private bool ShouldAllowNegativeSchool(int row, int score)
    {
        if (row > 2 || score >= 0)
            return false;

        return CurrentPlayer.School >= 6 && -score <= 3 && -score <= CurrentPlayer.School / 2;
    }

    private bool ShouldPreferCombinationsOverSchool()
    {
        var schoolCells = CountBusyMainCells(0, 6);
        var combinationCells = CountBusyMainCells(6, RowCount - 1);
        return schoolCells >= combinationCells + 6 || schoolCells >= 18 && combinationCells < 18;
    }

    private int CountBusyMainCells(int firstRow, int endRow)
    {
        var count = 0;
        for (var row = firstRow; row < endRow; row++)
            count += CountBusyMainCellsInRow(row);

        return count;
    }

    private int CountBusyMainCellsInRow(int row)
    {
        var count = 0;
        for (var column = 0; column < ColumnCount - 1; column++)
        {
            if (CurrentPlayer.Table[row, column] != EmptyCell)
                count++;
        }

        return count;
    }

    private int CountBusyMainCellsInColumn(int column)
    {
        var count = 0;
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.Table[row, column] != EmptyCell)
                count++;
        }

        return count;
    }

    private int CountCrossesInMainRow(int row)
    {
        var count = 0;
        for (var column = 0; column < ColumnCount - 1; column++)
        {
            if (CurrentPlayer.Table[row, column] == 0)
                count++;
        }

        return count;
    }

    private int CountCrossesInMainColumn(int column)
    {
        var count = 0;
        for (var row = 6; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.Table[row, column] == 0)
                count++;
        }

        return count;
    }

    private Dictionary<int, int> GetDiceCounts()
    {
        return Enumerable.Range(1, 6).ToDictionary(value => value, value => _dice.Count(die => die == value));
    }

    private int CountDice(int value) => _dice.Count(die => die == value);

    private bool HasFreeCombinationMainCell()
    {
        for (var row = 6; row < RowCount - 1; row++)
        {
            for (var column = 0; column < ColumnCount - 1; column++)
            {
                if (CurrentPlayer.Table[row, column] == EmptyCell)
                    return true;
            }
        }

        return false;
    }

    // Жесткие решения для готовых сильных ходов.
    // Этот блок срабатывает до общей оценки весов, чтобы ИИ не испортил очевидно сильную раздачу
    // и мог тактически закрыть/заблокировать приз даже не самым большим количеством очков.
    private int GetForcedPremiumMove()
    {
        var counts = GetDiceCounts();

        if (CurrentPlayer.GetFreeCell(7) != -1 && TwoPairsScore(counts) > 0 && IsTacticalPrizeMove(7))
            return 7;

        if (CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5))
            return 13;

        if (CurrentPlayer.GetFreeCell(12) != -1 && CurrentPlayer.School >= 6 && counts.Any(pair => pair.Value >= 4 && pair.Key >= 5))
            return 12;

        if (_rollCount != 1)
            return -1;

        if (CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4))
            return 12;

        if (CurrentPlayer.GetFreeCell(9) != -1 && FullHouseScore(counts) > 0)
            return 9;

        if (CurrentPlayer.GetFreeCell(11) != -1 && CalculateScore(11) > 0)
            return 11;

        if (CurrentPlayer.GetFreeCell(10) != -1 && CalculateScore(10) > 0)
            return 10;

        if (ShouldTakeOpeningSum(counts))
            return 14;

        if (CurrentPlayer.GetFreeCell(8) != -1 && counts.Any(pair => pair.Value == 3 && pair.Key >= 5))
            return 8;

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

    // После третьего броска ИИ может принудительно записать школу,
    // если нет более ценной готовой комбинации и ровно три кости нужного номинала уже собраны.
    private int GetForcedSchoolMove()
    {
        if (_rollCount < 3)
            return -1;

        var counts = GetDiceCounts();
        if (HasValuableFinishedCombination(counts))
            return -1;

        return counts
            .Where(pair => pair.Value == 3 && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1)
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

        var twoPairsScore = CalculateScore(7);
        if (twoPairsScore is <= 0 or >= 18 || IsTacticalPrizeMove(7))
            return -1;

        return CalculateScore(6) > 0 ? 6 : -1;
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
        if (CurrentPlayer.GetFreeCell(9) != -1 && FullHouseScore(counts) > 0)
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

    private bool IsGoodOpeningSum()
    {
        return _rollCount == 1 && _dice.Sum() >= 24;
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
            || CurrentPlayer.GetFreeCell(9) != -1 && FullHouseScore(counts) > 0
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

            var priority = count * 12 + value;
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

    // Ищет строку, где текущий ход может закрыть свой приз или помешать ближайшему призу соперника.
    private int GetUrgentComputerTargetRow()
    {
        for (var row = 6; row < RowCount - 1; row++)
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

        return CurrentPlayer.GetFreeCell(10) != -1 && smallStraight >= straightTarget
            || CurrentPlayer.GetFreeCell(11) != -1 && bigStraight >= straightTarget;
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
        if (count >= 3)
            return !preferCombinations || value >= 5;

        if (count < 2)
            return false;

        if (preferCombinations)
            return false;

        if (CurrentPlayer.School >= 0)
            return value >= 5 && CurrentPlayer.Difficulty == ComputerDifficulty.Careful;

        return true;
    }

    private readonly record struct ComputerMove(int Row, int Score, int AdjustedScore);
}
