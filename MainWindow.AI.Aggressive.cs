namespace AbacaWpf;

public partial class MainWindow
{
    private static readonly int[] AggressiveLightRows = [14, 7, 8, 6];
    private static readonly int[] AggressiveSchoolCoreRows = [5, 4, 3, 2];

    private bool IsAggressiveMode()
    {
        return CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive;
    }

    private int GetAggressiveMove()
    {
        if (!IsAggressiveMode())
            return -1;

        var openingMajorRow = GetAggressiveOpeningMajorMove();
        if (openingMajorRow >= 0)
            return openingMajorRow;

        var readyRareRow = GetAggressiveReadyRareMove();
        if (readyRareRow >= 0)
            return readyRareRow;

        var phaseOneRow = GetAggressivePhaseOneMove();
        if (phaseOneRow >= 0)
            return phaseOneRow;

        return GetAggressivePhaseTwoMove();
    }

    private bool ShouldStopOnAggressiveMove(ComputerMove best)
    {
        if (!IsAggressiveMode())
            return false;

        var aggressiveRow = GetAggressiveMove();
        if (aggressiveRow < 0 || aggressiveRow != best.Row || best.Score <= 0)
            return false;

        if (ShouldContinueAggressiveSchoolDevelopment(best))
            return false;

        return true;
    }

    private bool ApplyAggressiveKeepStrategy((int Value, int Count)[] counts)
    {
        if (!IsAggressiveMode())
            return false;

        if (IsAggressivePhaseOneActive())
            return ApplyAggressivePhaseOneKeepStrategy(counts);

        return ApplyAggressivePhaseTwoKeepStrategy(counts);
    }

    private bool IsAggressivePhaseOneActive()
    {
        if (!IsAggressiveMode())
            return false;

        return AggressiveLightRows.Any(row => CurrentPlayer.Table[row, ColumnCount - 1] == EmptyCell);
    }

    private bool IsAggressiveFirstTurn()
    {
        return CountBusyMainCells(0, RowCount - 1) == 0;
    }

    private int GetAggressiveOpeningMajorMove()
    {
        if (_rollCount != 1)
            return -1;

        var counts = GetDiceCounts();
        if (CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5))
            return 13;

        if (CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4))
            return 12;

        return -1;
    }

    private int GetAggressiveReadyRareMove()
    {
        if (!IsAggressiveMode())
            return -1;

        var rareRows = new[] { 13, 12, 11, 10, 9 };
        foreach (var row in rareRows)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
            if (score <= 0)
                continue;

            if (row == 9 && !IsFullHouseWorthRecording(GetDiceCounts()))
                continue;

            if (GetTacticalPrizeUrgency(row) >= 140 || !HasUrgentAggressiveLightRace())
                return row;
        }

        return -1;
    }

    private bool ShouldContinueAggressiveSchoolDevelopment(ComputerMove best)
    {
        if (_rollCount >= 3)
            return false;

        var counts = GetDiceCounts();
        var schoolValue = 0;
        if (best.Row is >= 0 and < 6 && counts.TryGetValue(best.Row + 1, out var schoolCount) && schoolCount >= 3)
            schoolValue = best.Row + 1;
        else if (best.Row == 8)
            schoolValue = counts
                .Where(pair => pair.Value >= 3 && CurrentPlayer.GetFreeCell(pair.Key - 1) != -1)
                .OrderByDescending(pair => pair.Value)
                .Select(pair => pair.Key)
                .FirstOrDefault();

        if (schoolValue == 0)
            return false;

        var schoolRow = schoolValue - 1;
        var schoolColumn = CurrentPlayer.GetFreeCell(schoolRow);
        var schoolIsImmediateTactical = CompletesOwnRowPrize(schoolRow)
            || schoolColumn >= 0 && schoolColumn < ColumnCount - 1 && CompletesOwnColumnPrize(schoolColumn)
            || schoolColumn >= 0 && schoolColumn < ColumnCount - 1 && ShouldBlockOpponentColumnPrize(schoolColumn);
        if (schoolIsImmediateTactical)
            return false;

        var canChaseAbaca = CurrentPlayer.GetFreeCell(13) != -1;
        var canChaseKare = CurrentPlayer.GetFreeCell(12) != -1;
        var canImproveSchool = CountDice(schoolValue) < DiceCount;

        return canChaseAbaca || canChaseKare || canImproveSchool;
    }

    private bool HasUrgentAggressiveLightRace()
    {
        foreach (var row in AggressiveLightRows)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1 || CurrentPlayer.Table[row, ColumnCount - 1] != EmptyCell)
                continue;

            var score = CalculateScore(row);
            if (score <= 0)
                continue;

            var urgency = GetTacticalPrizeUrgency(row);
            if (urgency >= 140)
                return true;

            var delta = GetAggressiveRaceDelta(row);
            if (delta < 0 && score >= GetAggressiveThresholdScore(row))
                return true;

            if (delta <= 0 && CountBusyMainCellsInRow(row) >= ColumnCount - 2)
                return true;
        }

        return false;
    }

    private int GetAggressivePhaseOneMove()
    {
        if (!IsAggressivePhaseOneActive())
            return -1;

        var counts = GetDiceCounts();
        if (_rollCount == 1
            && IsAggressiveFirstTurn()
            && !HasAggressiveOpeningMajorException(counts)
            && CurrentPlayer.GetFreeCell(14) != -1
            && CalculateScore(14) > 0)
        {
            return 14;
        }

        foreach (var row in AggressiveLightRows)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1 || CurrentPlayer.Table[row, ColumnCount - 1] != EmptyCell)
                continue;

            var score = CalculateScore(row);
            if (score <= 0)
                continue;

            var delta = GetAggressiveRaceDelta(row);
            if (delta > 0)
                continue;
            if (row == 6 && ShouldDeferAggressivePairForTripleDevelopment(counts, delta))
                continue;

            if (delta == 0)
                return row;

            if (score >= GetAggressiveThresholdScore(row))
                return row;
        }

        return -1;
    }

    private int GetAggressivePhaseTwoMove()
    {
        if (IsAggressivePhaseOneActive())
            return -1;

        var shouldProtectSchool = HasFragileSchoolEndgame()
            || HasHeavySchoolEndgamePressure()
            || ShouldProtectExistingLeadWithSchool();
        if (!shouldProtectSchool)
            return -1;

        var schoolSafetyCrossOutRow = GetAggressiveSchoolSafetyCrossOutRow();
        if (schoolSafetyCrossOutRow >= 0)
            return schoolSafetyCrossOutRow;

        foreach (var row in AggressiveSchoolCoreRows)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
            if (score > 0)
                return row;
        }

        return -1;
    }

    private int GetAggressiveSchoolSafetyCrossOutRow()
    {
        var affordableSchoolRow = GetAffordableNegativeSchoolRow();
        if (affordableSchoolRow < 0)
            return -1;

        var score = CalculateScore(affordableSchoolRow);
        var remainingSchool = CurrentPlayer.School + score;
        if (score > -4)
            return -1;
        if (remainingSchool > 6)
            return -1;

        var sacrificialRow = GetPreferredSchoolSafetyCrossOutRow(affordableSchoolRow);
        if (sacrificialRow is 10 or 11 or 13)
            return sacrificialRow;

        if (ShouldProtectExistingLeadWithSchool())
        {
            var emergencyRow = GetEmergencySacrificialCrossOutRow();
            if (emergencyRow >= 0)
                return emergencyRow;
        }

        return -1;
    }

    private bool ApplyAggressivePhaseOneKeepStrategy((int Value, int Count)[] counts)
    {
        var targetRow = GetAggressivePhaseOneTargetRow();
        if (targetRow < 0)
            return false;

        switch (targetRow)
        {
            case 14:
                return KeepAggressiveSumDice();
            case 7:
                return KeepAggressiveTwoPairsDice(counts);
            case 8:
                return KeepAggressiveTripleDice(counts);
            case 6:
                return KeepAggressivePairDice(counts);
            default:
                return false;
        }
    }

    private bool ApplyAggressivePhaseTwoKeepStrategy((int Value, int Count)[] counts)
    {
        var shouldProtectSchool = HasFragileSchoolEndgame()
            || HasHeavySchoolEndgamePressure()
            || ShouldProtectExistingLeadWithSchool();

        if (shouldProtectSchool)
        {
            var schoolCore = counts
                .Where(item => item.Value is >= 3 and <= 6 && item.Count > 0 && CurrentPlayer.GetFreeCell(item.Value - 1) != -1)
                .OrderByDescending(item => GetCriticalSchoolKeepPriority(item.Value, item.Count))
                .ThenByDescending(item => item.Value)
                .FirstOrDefault();

            if (schoolCore.Count > 0)
            {
                KeepValues([schoolCore.Value]);
                return true;
            }

            var hasOpenCoreSchool = AggressiveSchoolCoreRows.Any(row => CurrentPlayer.GetFreeCell(row) != -1);
            if (hasOpenCoreSchool)
            {
                Array.Fill(_fixedDice, false);
                return true;
            }
        }

        return KeepAggressiveLateFullHousePlan(counts);
    }

    private int GetAggressivePhaseOneTargetRow()
    {
        foreach (var row in AggressiveLightRows)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1 || CurrentPlayer.Table[row, ColumnCount - 1] != EmptyCell)
                continue;

            if (GetAggressiveRaceDelta(row) <= 0)
                return row;
        }

        return -1;
    }

    private bool HasAggressiveOpeningMajorException(Dictionary<int, int> counts)
    {
        return CurrentPlayer.GetFreeCell(13) != -1 && counts.Any(pair => pair.Value == 5)
            || CurrentPlayer.GetFreeCell(12) != -1 && counts.Any(pair => pair.Value >= 4);
    }

    private int GetAggressiveRaceDelta(int row)
    {
        var opponent = _players[1 - _currentPlayerIndex];
        return CountBusyMainCellsInRow(row) - CountBusyMainCellsInRow(opponent, row);
    }

    private int CountBusyMainCellsInRow(Player player, int row)
    {
        var count = 0;
        for (var column = 0; column < ColumnCount - 1; column++)
        {
            if (player.Table[row, column] != EmptyCell)
                count++;
        }

        return count;
    }

    private int GetAggressiveThresholdScore(int row)
    {
        var opening = _rollCount == 1;
        return row switch
        {
            14 => opening ? 40 : 20,
            7 => opening ? 32 : 16,
            8 => opening ? 24 : 12,
            6 => opening ? 16 : 8,
            _ => int.MaxValue
        };
    }

    private bool ShouldDeferAggressivePairForTripleDevelopment(Dictionary<int, int> counts, int raceDelta)
    {
        if (CurrentPlayer.GetFreeCell(6) == -1)
            return false;
        if (!counts.Any(pair => pair.Value >= 3 && CanDevelopAggressiveTriple(pair.Key)))
            return false;
        if (IsTacticalPrizeMove(6))
            return false;

        return raceDelta < 0;
    }

    private bool KeepAggressiveSumDice()
    {
        var values = _rollCount == 1
            ? _dice.Where(value => value >= 5).ToHashSet()
            : _dice.Where(value => value >= 4).ToHashSet();

        if (values.Count == 0)
        {
            Array.Fill(_fixedDice, false);
            return true;
        }

        KeepValues(values);
        return true;
    }

    private bool KeepAggressiveTwoPairsDice((int Value, int Count)[] counts)
    {
        var pairValues = counts
            .Where(item => item.Count >= 2)
            .Select(item => item.Value)
            .ToHashSet();

        if (pairValues.Count == 0)
        {
            Array.Fill(_fixedDice, false);
            return true;
        }

        KeepValues(pairValues);
        return true;
    }

    private bool KeepAggressiveTripleDice((int Value, int Count)[] counts)
    {
        var target = counts
            .Where(item => item.Count >= 2 && CanDevelopAggressiveTriple(item.Value))
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();

        if (target.Count == 0)
        {
            Array.Fill(_fixedDice, false);
            return true;
        }

        KeepValues([target.Value]);
        return true;
    }

    private bool KeepAggressivePairDice((int Value, int Count)[] counts)
    {
        var target = counts
            .Where(item => item.Count >= 2)
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .FirstOrDefault();

        if (target.Count == 0)
        {
            Array.Fill(_fixedDice, false);
            return true;
        }

        KeepValues([target.Value]);
        return true;
    }

    private bool CanDevelopAggressiveTriple(int value)
    {
        return CurrentPlayer.GetFreeCell(value - 1) != -1
            || CurrentPlayer.GetFreeCell(8) != -1
            || CurrentPlayer.GetFreeCell(12) != -1
            || CurrentPlayer.GetFreeCell(13) != -1
            || CurrentPlayer.GetFreeCell(9) != -1;
    }

    private bool KeepAggressiveLateFullHousePlan((int Value, int Count)[] counts)
    {
        if (HasFreeCombinationMainCell()
            && Enumerable.Range(6, RowCount - 7).Any(row => CurrentPlayer.GetFreeCell(row) != -1 && row is not (9 or 10 or 11 or 13)))
        {
            return false;
        }

        if (CurrentPlayer.GetFreeCell(9) == -1)
            return false;

        var pairValues = counts
            .Where(item => item.Count >= 2)
            .Select(item => item.Value)
            .ToHashSet();

        if (pairValues.Count >= 2)
        {
            KeepValues(pairValues);
            return true;
        }

        if (pairValues.Count == 1)
        {
            KeepValues(pairValues);
            return true;
        }

        return false;
    }
}
