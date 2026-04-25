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
                if (score == 0 && count >= 3)
                    priority += GetSchoolZeroBonus(row);
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
            priority += ShouldBlockOpponentRowPrize(row) ? 72 : OpponentPrizeIsOpen(row, ColumnCount - 1) ? 34 : 16;
        if (column >= 0 && column < ColumnCount - 1 && CompletesOwnColumnPrize(column))
            priority += ShouldBlockOpponentColumnPrize(column) ? 64 : OpponentPrizeIsOpen(RowCount - 1, column) ? 30 : 14;
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

    // Школьный ноль - хороший результат: строка закрывается без потери школьного запаса.
    // Это особенно важно, когда школа отстает от комбинаций или в этой строке еще мало записей.
    private int GetSchoolZeroBonus(int row)
    {
        var bonus = 28;
        var schoolFill = CountBusyMainCells(0, 6);
        var combinationFill = CountBusyMainCells(6, RowCount - 1);
        var rowBusy = CountBusyMainCellsInRow(row);

        if (schoolFill + 4 < combinationFill)
            bonus += 26;
        if (rowBusy <= 1)
            bonus += 16;
        if (IsSchoolPrizeRaceTarget(row))
            bonus += 30;

        return bonus;
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
            7 => 28,  // две пары
            6 => 30,  // пара
            8 => 34,  // тройку тоже не вычеркиваем рано
            9 => 40,  // фул лучше беречь
            12 => 46, // каре лучше беречь
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
        return IsRowOpenForCrossOut(13)
            || IsRowOpenForCrossOut(11)
            || IsRowOpenForCrossOut(10);
    }

    private bool IsRowOpenForCrossOut(int row)
    {
        return CurrentPlayer.GetFreeCell(row) != -1 && CalculateScore(row) == 0;
    }

    private bool ShouldAllowNegativeSchool(int row, int score)
    {
        if (score >= 0)
            return false;

        if (!IsLegalSchoolMove(row, score))
            return false;

        if (score == -1)
            return true;

        if (!IsNegativeSchoolRiskAcceptable(row, score))
            return false;

        if (CountDice(row + 1) > 0 && IsLastOpenSchoolCell(row))
            return true;

        if (CountDice(row + 1) > 0 && CanSafelyCloseRemainingSchool(row, score))
            return true;

        return CurrentPlayer.School > 0 && -score <= 6 && -score <= CurrentPlayer.School;
    }

    private bool IsLegalSchoolMove(int row, int score)
    {
        if (row is < 0 or >= 6)
            return false;

        if (score >= 0)
            return true;

        if (CountDice(row + 1) == 0 && HasFreeCombinationMainCell())
            return false;

        if (-score > CurrentPlayer.School && !CurrentPlayer.IsOnlySchoolFree())
            return false;

        return true;
    }

    private bool IsLegalComputerMove(int row, int score)
    {
        if (row is < 0 or >= RowCount - 1 || CurrentPlayer.GetFreeCell(row) == -1)
            return false;

        if (row >= 6)
            return true;

        if (!IsLegalSchoolMove(row, score))
            return false;

        return score >= 0 || !HasFreeCombinationMainCell() || ShouldAllowNegativeSchool(row, score);
    }

    private bool IsNegativeSchoolRiskAcceptable(int row, int score)
    {
        if (row is < 0 or >= 6 || score >= 0)
            return false;

        var remainingSchool = CurrentPlayer.School + score;
        if (remainingSchool < 0)
            return false;

        var highDemand = Enumerable.Range(1, 6)
            .Where(value => value >= 5 && value - 1 != row && CurrentPlayer.GetFreeCell(value - 1) != -1)
            .Sum(value => (ColumnCount - 1 - CountBusyMainCellsInRow(value - 1)) * value);
        if (highDemand > 0 && remainingSchool < highDemand)
            return false;

        var maxRemainingRowCost = Enumerable.Range(0, 6)
            .Where(schoolRow => schoolRow != row && CurrentPlayer.GetFreeCell(schoolRow) != -1)
            .Select(GetMinimumSchoolCloseCost)
            .DefaultIfEmpty(0)
            .Max();

        return remainingSchool >= maxRemainingRowCost;
    }

    private bool IsLastOpenSchoolCell(int row)
    {
        if (row is < 0 or >= 6 || CurrentPlayer.GetFreeCell(row) == -1)
            return false;

        var freeSchoolCells = 0;
        for (var schoolRow = 0; schoolRow < 6; schoolRow++)
        {
            for (var column = 0; column < ColumnCount - 1; column++)
            {
                if (CurrentPlayer.Table[schoolRow, column] == EmptyCell)
                    freeSchoolCells++;
            }
        }

        return freeSchoolCells == 1;
    }

    private bool CanSafelyCloseRemainingSchool(int targetRow, int targetScore)
    {
        if (targetRow is < 0 or >= 6 || targetScore >= 0)
            return false;

        var remainingDemand = 0;
        for (var row = 0; row < 6; row++)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            if (row == targetRow)
                continue;

            remainingDemand += GetMinimumSchoolCloseCost(row);
        }

        return CurrentPlayer.School + targetScore >= remainingDemand;
    }

    private int GetMinimumSchoolCloseCost(int row)
    {
        if (row is < 0 or >= 6 || CurrentPlayer.GetFreeCell(row) == -1)
            return 0;

        var freeCells = 0;
        for (var column = 0; column < ColumnCount - 1; column++)
        {
            if (CurrentPlayer.Table[row, column] == EmptyCell)
                freeCells++;
        }

        return freeCells * (row + 1);
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
