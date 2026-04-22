namespace AbacaWpf;

public partial class MainWindow
{
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
        if (row == 12 && score > 0 && GetKareValue(GetDiceCounts()) < 5)
            bonus -= CountBusyMainCellsInRow(row) >= 2 || CurrentPlayer.School <= 6 ? 90 : 35;
        if (row == 9 && score > 0 && FullHouseScore(GetDiceCounts()) < 22)
            bonus -= CountBusyMainCellsInRow(row) >= 2 || CurrentPlayer.School <= 6 ? 70 : 28;

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

        return CurrentPlayer.School > 0 && -score <= 3 && -score <= CurrentPlayer.School;
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

}
