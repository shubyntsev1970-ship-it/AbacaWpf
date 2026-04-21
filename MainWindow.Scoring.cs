namespace AbacaWpf;

public partial class MainWindow
{
    // Базовый расчет очков по текущим кубикам.
    // Здесь собраны только правила комбинаций: школа, пара, две пары, фул, стриты, каре, абак и сумма.
    private int CalculateScore(int row)
    {
        if (row < 6)
        {
            var number = row + 1;
            return number * _dice.Count(die => die == number) - number * 3;
        }

        var counts = Enumerable.Range(1, 6).ToDictionary(value => value, value => _dice.Count(die => die == value));
        var score = row switch
        {
            6 => counts.Where(pair => pair.Value >= 2).Select(pair => pair.Key * 2).DefaultIfEmpty(0).Max(),
            7 => TwoPairsScore(counts),
            8 => counts.Where(pair => pair.Value >= 3).Select(pair => pair.Key * 3).DefaultIfEmpty(0).Max(),
            9 => FullHouseScore(counts),
            10 => Enumerable.Range(1, 5).All(value => counts[value] == 1) ? 15 : 0,
            11 => Enumerable.Range(2, 5).All(value => counts[value] == 1) ? 20 : 0,
            12 => counts.Where(pair => pair.Value >= 4).Select(pair => pair.Key * 4 + 20).DefaultIfEmpty(0).Max(),
            13 => counts.Where(pair => pair.Value == 5).Select(pair => pair.Key * 5 + 50).DefaultIfEmpty(0).Max(),
            14 => _dice.Sum(),
            _ => 0
        };

        return _rollCount == 1 ? score * 2 : score;
    }

    // Две пары могут быть составлены и из 4-5 одинаковых костей:
    // например 55555 подходит как пара, две пары, тройка, фул, каре и абак.
    private static int TwoPairsScore(Dictionary<int, int> counts)
    {
        var pairs = GetTwoPairValues(counts);
        return pairs.Length == 2 ? pairs.Sum() * 2 : 0;
    }

    private static int[] GetTwoPairValues(Dictionary<int, int> counts)
    {
        return counts
            .SelectMany(pair => Enumerable.Repeat(pair.Key, pair.Value / 2))
            .OrderByDescending(value => value)
            .Take(2)
            .ToArray();
    }

    private static int FullHouseScore(Dictionary<int, int> counts)
    {
        var five = counts.FirstOrDefault(pair => pair.Value == 5);
        if (five.Key > 0)
            return five.Key * 5;

        var triple = counts.Where(pair => pair.Value == 3).Select(pair => pair.Key).DefaultIfEmpty(0).Max();
        var pairValue = counts.Where(pair => pair.Value == 2).Select(pair => pair.Key).DefaultIfEmpty(0).Max();
        return triple > 0 && pairValue > 0 ? triple * 3 + pairValue * 2 : 0;
    }
}
