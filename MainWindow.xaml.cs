using System.IO;
using System.Media;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace AbacaWpf;

public partial class MainWindow : Window
{
    // Board geometry and table markers.
    private const int DiceCount = 5;
    private const int PlayerCount = 2;
    private const int RowCount = 16;
    private const int ColumnCount = 6;
    private const int EmptyCell = -100;
    private const double DiceSize = 100;
    private const double DiceHorizontalGap = DiceSize / 2;

    // Runtime game state and UI cell references.
    private readonly Random _random = new();
    private readonly DispatcherTimer _rollTimer;
    private readonly Player[] _players = [new(), new()];
    private readonly int[] _dice = [1, 2, 3, 4, 5];
    private readonly bool[] _fixedDice = new bool[DiceCount];
    private readonly bool[] _selectedDice = new bool[DiceCount];
    private readonly TextBlock[,,] _cells = new TextBlock[PlayerCount, RowCount, ColumnCount];
    private readonly SoundPlayer _diceSoundPlayer = new(CreateDiceRollSound());
    private Border? _lastMoveBorder;
    private int _lastMoveRow = -1;
    private readonly string[] _rowLabels =
    [
        "1", "2", "3", "4", "5", "6", "пара", "пары", "тройка", "фул", "Мстр", "Бстр", "каре", "абак", "сумма", "приз"
    ];
    private readonly string[] _rowHotkeys =
    [
        "1", "2", "3", "4", "5", "6", "P", "D", "T", "F", "E", "S", "C", "A", "+", ""
    ];

    private int _currentPlayerIndex;
    private int _rollCount;
    private bool _isRolling;
    private bool _computerTurnInProgress;

    // Window startup and initial control construction.
    public MainWindow()
    {
        InitializeComponent();
        _rollTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(90) };
        _rollTimer.Tick += RollTimer_Tick;
        BuildTables();
        BuildCombinationButtons();
        BuildDice();
        Loaded += (_, _) =>
        {
            StartNewGame();
            Focus();
        };
    }

    private Player CurrentPlayer => _players[_currentPlayerIndex];

    // Game setup: asks for players and resets all state before the first turn.
    private void StartNewGame()
    {
        var dialog = new StartGameWindow(_random) { Owner = this };
        if (dialog.ShowDialog() != true)
        {
            Close();
            return;
        }

        _computerTurnInProgress = false;
        _players[0] = new Player(dialog.FirstPlayerName);
        _players[1] = new Player(dialog.SecondPlayerName, dialog.PlayAgainstComputer, dialog.ComputerDifficulty);
        if (dialog.SecondPlayerStarts)
        {
            (_players[0], _players[1]) = (_players[1], _players[0]);
        }

        _currentPlayerIndex = 1;
        ClearTableValues();
        StartNextTurn();
    }

    // Table construction and cell rendering.
    private void BuildTables()
    {
        BuildPlayerTable(PlayerOneTable, 0);
        BuildPlayerTable(PlayerTwoTable, 1);
    }

    private void BuildPlayerTable(Grid grid, int playerIndex)
    {
        grid.Children.Clear();
        grid.RowDefinitions.Clear();
        grid.ColumnDefinitions.Clear();

        for (var row = 0; row < RowCount + 1; row++)
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star), MinHeight = 20 });

        for (var column = 0; column < ColumnCount + 1; column++)
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(column == 0 ? 1.55 : 0.375, GridUnitType.Star) });

        AddHeaderCell(grid, "", 0, 0);
        for (var column = 1; column <= 5; column++)
            AddHeaderCell(grid, column.ToString(), 0, column);
        AddHeaderCell(grid, "приз", 0, 6);

        for (var row = 0; row < RowCount; row++)
        {
            AddHeaderCell(grid, GetRowCaption(row), row + 1, 0);
            for (var column = 0; column < ColumnCount; column++)
            {
                var cell = AddValueCell(grid, row + 1, column + 1);
                _cells[playerIndex, row, column] = cell;
            }
        }
    }

    private static void AddHeaderCell(Grid grid, string text, int row, int column)
    {
        var border = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(118, 144, 160)),
            BorderThickness = new Thickness(1),
            Background = row == 0 || column == 0 ? new SolidColorBrush(Color.FromRgb(33, 78, 114)) : Brushes.Transparent
        };
        var fontSize = column == 0 && row > 0 ? 22 : text.Length > 7 ? 12 : text.Length > 4 ? 14 : 16;
        var label = CreateCaptionTextBlock(text, fontSize);
        border.Child = label;
        Grid.SetRow(border, row);
        Grid.SetColumn(border, column);
        grid.Children.Add(border);
    }

    private static TextBlock AddValueCell(Grid grid, int row, int column)
    {
        var border = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(138, 160, 174)),
            BorderThickness = new Thickness(1),
            Background = row >= 7 && row <= 15 ? Brushes.White : new SolidColorBrush(Color.FromRgb(247, 251, 253))
        };
        var content = new Grid();
        var label = new TextBlock
        {
            Foreground = new SolidColorBrush(Color.FromRgb(18, 26, 34)),
            FontWeight = FontWeights.Bold,
            FontSize = 21,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            TextAlignment = TextAlignment.Center
        };
        var crossMark = new System.Windows.Shapes.Path
        {
            Data = new GeometryGroup
            {
                Children = new GeometryCollection
                {
                    new LineGeometry(new Point(0, 0), new Point(1, 1)),
                    new LineGeometry(new Point(1, 0), new Point(0, 1))
                }
            },
            Stretch = Stretch.Fill,
            Stroke = new SolidColorBrush(Color.FromRgb(80, 92, 104)),
            StrokeThickness = 4,
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round,
            Margin = new Thickness(2),
            IsHitTestVisible = false,
            Visibility = Visibility.Collapsed
        };
        content.Children.Add(label);
        content.Children.Add(crossMark);
        border.Child = content;
        Grid.SetRow(border, row);
        Grid.SetColumn(border, column);
        grid.Children.Add(border);
        return label;
    }

    // Bottom combination buttons and dice visuals.
    private void BuildCombinationButtons()
    {
        CombinationButtons.Items.Clear();
        for (var row = 0; row < RowCount - 1; row++)
        {
            var button = new Button
            {
                Content = CreateCaptionTextBlock(GetRowCaption(row), 18),
                Tag = row,
                Style = (Style)FindResource("CommandButton"),
                FontSize = 18,
                ToolTip = GetCombinationHelp(row)
            };
            button.Click += CombinationButton_Click;
            CombinationButtons.Items.Add(button);
        }
    }

    private void BuildDice()
    {
        DiceItems.Items.Clear();
        for (var i = 0; i < DiceCount; i++)
        {
            var index = i;
            var transform = new TransformGroup();
            transform.Children.Add(new ScaleTransform(1, 1));
            transform.Children.Add(new RotateTransform(0));

            var border = new Border
            {
                Width = DiceSize,
                Height = DiceSize,
                Margin = new Thickness(DiceHorizontalGap / 2, _fixedDice[index] ? 0 : DiceSize, DiceHorizontalGap / 2, _fixedDice[index] ? DiceSize : 0),
                CornerRadius = new CornerRadius(8),
                BorderThickness = new Thickness(_fixedDice[index] ? _selectedDice[index] ? 6 : 0.625 : _selectedDice[index] ? 9 : 2),
                BorderBrush = _fixedDice[index]
                    ? new SolidColorBrush(Color.FromRgb(239, 68, 68))
                    : _selectedDice[index]
                        ? new SolidColorBrush(Color.FromRgb(34, 197, 94))
                        : new SolidColorBrush(Color.FromRgb(187, 208, 221)),
                Background = _fixedDice[index]
                    ? new LinearGradientBrush(Color.FromRgb(255, 99, 99), Color.FromRgb(185, 28, 28), 45)
                    : new LinearGradientBrush(Color.FromRgb(255, 255, 255), Color.FromRgb(219, 229, 237), 45),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius = 16,
                    ShadowDepth = 4,
                    Opacity = 0.28
                },
                RenderTransform = transform,
                RenderTransformOrigin = new Point(0.5, 0.5),
                Child = CreateDiceFace(_dice[index])
            };
            if (_isRolling && !_fixedDice[index])
                AnimateRollingDie(border);
            border.MouseLeftButtonDown += (_, _) => ToggleDie(index);
            DiceItems.Items.Add(border);
        }
    }

    private string GetRowCaption(int row) =>
        string.IsNullOrWhiteSpace(_rowHotkeys[row])
            ? _rowLabels[row]
            : $"{_rowLabels[row]} ({_rowHotkeys[row]})";

    private static TextBlock CreateCaptionTextBlock(string caption, double fontSize)
    {
        var label = new TextBlock
        {
            Foreground = Brushes.White,
            FontWeight = FontWeights.Bold,
            FontSize = fontSize,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            TextAlignment = TextAlignment.Center
        };

        var hotkeyStart = caption.LastIndexOf(" (", StringComparison.Ordinal);
        if (hotkeyStart < 0 || !caption.EndsWith(')'))
        {
            label.Text = caption;
            return label;
        }

        label.Inlines.Add(new Run(caption[..hotkeyStart]));
        label.Inlines.Add(new Run(caption[hotkeyStart..])
        {
            Foreground = new SolidColorBrush(Color.FromRgb(255, 99, 99)),
            FontWeight = FontWeights.Black
        });
        return label;
    }

    // Dice rendering, animation, and mouse selection.
    private static Image CreateDiceFace(int value)
    {
        return new Image
        {
            Source = new BitmapImage(new Uri($"pack://application:,,,/Assets/Dice_{value}.jpg", UriKind.Absolute)),
            Stretch = Stretch.Uniform,
            Margin = new Thickness(8),
            SnapsToDevicePixels = true
        };
    }

    private static void AnimateRollingDie(Border border)
    {
        var rotate = new DoubleAnimation(0, 18, TimeSpan.FromMilliseconds(90))
        {
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever
        };
        var scale = new DoubleAnimation(0.94, 1.04, TimeSpan.FromMilliseconds(90))
        {
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever
        };
        if (border.RenderTransform is TransformGroup group)
        {
            group.Children[0].BeginAnimation(ScaleTransform.ScaleXProperty, scale);
            group.Children[0].BeginAnimation(ScaleTransform.ScaleYProperty, scale);
            group.Children[1].BeginAnimation(RotateTransform.AngleProperty, rotate);
        }
    }

    // Turn flow: starts/stops rolling and prepares the next player.
    private void StartNextTurn()
    {
        UpdateScorePanel();
        _currentPlayerIndex = _currentPlayerIndex == 0 ? 1 : 0;
        _rollCount = 0;
        for (var i = 0; i < DiceCount; i++)
        {
            _fixedDice[i] = false;
            _selectedDice[i] = false;
        }

        _isRolling = true;
        RollButton.Content = "СТОП";
        _rollTimer.Start();
        PlayDiceRollSound();
        UpdateScorePanel();
        BuildDice();
        UpdateInputState();
        if (CurrentPlayer.IsComputer)
            BeginComputerTurn();
    }

    private void ToggleRoll()
    {
        if (_computerTurnInProgress)
            return;

        ToggleRollInternal();
    }

    private void ToggleRollInternal()
    {
        if (_rollCount >= 3 || _fixedDice.All(fixedDie => fixedDie))
            return;

        if (_isRolling)
            StopRolling();
        else
            StartRolling();

        UpdateScorePanel();
        BuildDice();
    }

    private void StartRolling()
    {
        _isRolling = true;
        RollButton.Content = "СТОП";
        _rollTimer.Start();
        PlayDiceRollSound();
    }

    private void StopRolling()
    {
        _isRolling = false;
        _rollTimer.Stop();
        _rollCount++;
        if (!_selectedDice.Any(isSelected => isSelected))
            _selectedDice[0] = true;
        RollButton.Content = "СТАРТ";
    }

    private void ToggleDie(int index)
    {
        if (_isRolling || _computerTurnInProgress)
            return;

        _fixedDice[index] = !_fixedDice[index];
        for (var i = 0; i < DiceCount; i++)
            _selectedDice[i] = i == index;
        BuildDice();
    }

    // Combination validation, scoring rules, and prize resolution.
    private void ScoreCombination(int row, bool fromComputer = false)
    {
        if (_computerTurnInProgress && !fromComputer)
            return;

        if (_isRolling)
            return;

        var column = CurrentPlayer.GetFreeCell(row);
        if (column == -1)
            return;

        var score = CalculateScore(row);
        if (row < 6 && score < 0 && -score > CurrentPlayer.School && !CurrentPlayer.IsOnlySchoolFree())
        {
            ShowLargeMessage("ABACA", "Недостаточно очков школы для такой записи. Выберите другую комбинацию.");
            return;
        }

        WriteTable(row, column, score);
        if (!IsGameOver())
            StartNextTurn();
    }

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

    private static int TwoPairsScore(Dictionary<int, int> counts)
    {
        var pairs = counts
            .SelectMany(pair => Enumerable.Repeat(pair.Key, pair.Value / 2))
            .OrderByDescending(value => value)
            .Take(2)
            .ToArray();
        return pairs.Length == 2 ? pairs.Sum() * 2 : 0;
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

    private void WriteTable(int row, int column, int score)
    {
        if (row < 6)
        {
            if (score > 0 && CurrentPlayer.School < 0)
            {
                CurrentPlayer.Score += score >= -CurrentPlayer.School ? -CurrentPlayer.School * 100 : score * 100;
            }
            else if (score < 0 && CurrentPlayer.School <= 0)
            {
                CurrentPlayer.Score += score * 100;
            }
            else if (score < 0 && CurrentPlayer.School > 0 && -score > CurrentPlayer.School)
            {
                CurrentPlayer.Score += (score + CurrentPlayer.School) * 100;
            }

            CurrentPlayer.School += score;
        }

        CurrentPlayer.SetBusyCell(row, column, score);
        CurrentPlayer.Score += score;
        WriteCell(_currentPlayerIndex, row, column, score);
        HighlightLastMove(_currentPlayerIndex, row, column);

        if (score == 0)
        {
            if (row >= 6)
            {
                CrossPrizeCellIfFree(CurrentPlayer, _currentPlayerIndex, row, ColumnCount - 1);
                CrossPrizeCellIfFree(CurrentPlayer, _currentPlayerIndex, RowCount - 1, column);
            }
        }
        else if (score < 0)
        {
            CrossPrizeCellIfFree(CurrentPlayer, _currentPlayerIndex, row, ColumnCount - 1);
        }

        AwardRowPrize(row);
        AwardColumnPrize(row, column);
        UpdateScorePanel();
    }

    private void AwardRowPrize(int row)
    {
        if (!CurrentPlayer.IsAllRowMainCellsBusy(row))
            return;

        if (CurrentPlayer.Table[row, ColumnCount - 1] == EmptyCell)
        {
            var value = row >= 6 ? CurrentPlayer.GetMaxInRow(row) : (row + 1) * 5;
            CurrentPlayer.SetBusyCell(row, ColumnCount - 1, value);
            CurrentPlayer.Score += value;
            WriteCell(_currentPlayerIndex, row, ColumnCount - 1, value);
        }

        var opponentIndex = 1 - _currentPlayerIndex;
        var opponent = _players[opponentIndex];
        CrossPrizeCellIfFree(opponent, opponentIndex, row, ColumnCount - 1);
    }

    private void AwardColumnPrize(int row, int column)
    {
        if (!CurrentPlayer.IsAllColumnBusy(column) || CurrentPlayer.Table[RowCount - 1, column] != EmptyCell)
            return;

        var value = CurrentPlayer.GetMaxInColumn(column);
        CurrentPlayer.SetBusyCell(RowCount - 1, column, value);
        CurrentPlayer.Score += value;
        WriteCell(_currentPlayerIndex, RowCount - 1, column, value);

        if (column < ColumnCount - 2)
        {
            var opponentIndex = 1 - _currentPlayerIndex;
            var opponent = _players[opponentIndex];
            CrossPrizeCellIfFree(opponent, opponentIndex, RowCount - 1, column);
        }
    }

    private void CrossPrizeCellIfFree(Player player, int playerIndex, int row, int column)
    {
        if (player.Table[row, column] != EmptyCell)
            return;

        player.SetBusyCell(row, column, 0);
        WriteCell(playerIndex, row, column, 0);
    }

    // Table write effects: numbers, crosses, and last-move highlight.
    private void WriteCell(int playerIndex, int row, int column, int score)
    {
        var cell = _cells[playerIndex, row, column];
        cell.Text = score == 0 ? "" : score.ToString();
        SetCrossMarkVisibility(cell, score == 0);
        cell.Foreground = score < 0
            ? new SolidColorBrush(Color.FromRgb(190, 65, 65))
            : score == 0
                ? new SolidColorBrush(Color.FromRgb(80, 92, 104))
                : new SolidColorBrush(Color.FromRgb(18, 26, 34));
        cell.FontSize = score >= 100 || score < -9 ? 19 : 23;
        cell.FontWeight = FontWeights.Bold;
        AnimateCellWrite(cell, score);
    }

    private static void SetCrossMarkVisibility(TextBlock cell, bool isVisible)
    {
        if (cell.Parent is not Grid content)
            return;

        foreach (var child in content.Children)
        {
            if (child is System.Windows.Shapes.Path path)
            {
                path.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;
                return;
            }
        }
    }

    private void HighlightLastMove(int playerIndex, int row, int column)
    {
        if (_lastMoveBorder is not null && _lastMoveRow >= 0)
        {
            ResetValueCellStyle(_lastMoveBorder, _lastMoveRow);
        }

        var border = GetCellBorder(_cells[playerIndex, row, column]);
        if (border is null)
            return;

        _lastMoveBorder = border;
        _lastMoveRow = row;
        border.Background = new SolidColorBrush(Color.FromRgb(255, 237, 213));
        border.BorderBrush = new SolidColorBrush(Color.FromRgb(220, 38, 38));
        border.BorderThickness = new Thickness(2);
    }

    private static void ResetValueCellStyle(Border border, int row)
    {
        border.Background = row >= 6 && row <= 14
            ? Brushes.White
            : new SolidColorBrush(Color.FromRgb(247, 251, 253));
        border.BorderBrush = new SolidColorBrush(Color.FromRgb(138, 160, 174));
        border.BorderThickness = new Thickness(1);
    }

    private static Border? GetCellBorder(TextBlock cell)
    {
        if (cell.Parent is Grid { Parent: Border border })
            return border;

        return cell.Parent as Border;
    }

    private static void AnimateCellWrite(TextBlock cell, int score)
    {
        cell.RenderTransformOrigin = new Point(0.5, 0.5);
        cell.RenderTransform = new ScaleTransform(0.65, 0.65);

        var scale = new DoubleAnimation(0.65, 1, TimeSpan.FromMilliseconds(260))
        {
            EasingFunction = new BackEase { Amplitude = 0.25, EasingMode = EasingMode.EaseOut }
        };
        var opacity = new DoubleAnimation(0.15, 1, TimeSpan.FromMilliseconds(220));

        cell.BeginAnimation(OpacityProperty, opacity);
        if (cell.RenderTransform is ScaleTransform transform)
        {
            transform.BeginAnimation(ScaleTransform.ScaleXProperty, scale);
            transform.BeginAnimation(ScaleTransform.ScaleYProperty, scale);
        }

        var border = GetCellBorder(cell);
        if (border is not null)
        {
            var flash = score < 0
                ? Color.FromRgb(255, 219, 219)
                : score == 0
                    ? Color.FromRgb(229, 235, 240)
                    : Color.FromRgb(255, 245, 196);
            var brush = new SolidColorBrush(flash);
            border.Background = brush;
            brush.BeginAnimation(SolidColorBrush.ColorProperty,
                new ColorAnimation(flash, Color.FromRgb(255, 255, 255), TimeSpan.FromMilliseconds(520)));
        }
    }

    // Score panel, enabled/disabled controls, and end-of-game messaging.
    private bool IsGameOver()
    {
        if (_currentPlayerIndex != 1 || !CurrentPlayer.IsAllColumnBusy(ColumnCount - 2))
            return false;

        var message = _players[0].Score == _players[1].Score
            ? "НИЧЬЯ!"
            : _players[0].Score > _players[1].Score
                ? $"Победил {_players[0].Name}: +{_players[0].Score - _players[1].Score} очков"
                : $"Победил {_players[1].Name}: +{_players[1].Score - _players[0].Score} очков";

        UpdateScorePanel();
        ShowLargeMessage("Игра закончена", message);
        return true;
    }

    private void UpdateScorePanel()
    {
        PlayerOneNameText.Text = _players[0].Name;
        PlayerTwoNameText.Text = _players[1].Name;
        PlayerOneScoreText.Text = _players[0].Score.ToString();
        PlayerTwoScoreText.Text = _players[1].Score.ToString();
        PlayerOneSchoolText.Text = _players[0].School.ToString();
        PlayerTwoSchoolText.Text = _players[1].School.ToString();
        TurnText.Text = $"Ход: {_players[_currentPlayerIndex].Name}";
        RollCountText.Text = $"Бросок: {_rollCount} / 3";
        PlayerOnePanel.BorderBrush = _currentPlayerIndex == 0 ? new SolidColorBrush(Color.FromRgb(246, 211, 101)) : new SolidColorBrush(Color.FromRgb(62, 97, 124));
        PlayerTwoPanel.BorderBrush = _currentPlayerIndex == 1 ? new SolidColorBrush(Color.FromRgb(246, 211, 101)) : new SolidColorBrush(Color.FromRgb(62, 97, 124));
        PlayerOnePanel.BorderThickness = _currentPlayerIndex == 0 ? new Thickness(3) : new Thickness(1);
        PlayerTwoPanel.BorderThickness = _currentPlayerIndex == 1 ? new Thickness(3) : new Thickness(1);
    }

    private void UpdateInputState()
    {
        var humanCanAct = !CurrentPlayer.IsComputer && !_computerTurnInProgress;
        RollButton.IsEnabled = humanCanAct;
        foreach (var item in CombinationButtons.Items)
        {
            if (item is Button button)
            {
                button.IsEnabled = true;
                button.IsHitTestVisible = humanCanAct;
                button.Opacity = humanCanAct ? 1 : 0.72;
            }
        }
    }

    private void ClearTableValues()
    {
        foreach (var cell in _cells)
        {
            cell.Text = "";
            SetCrossMarkVisibility(cell, false);
        }

        if (_lastMoveBorder is not null && _lastMoveRow >= 0)
        {
            ResetValueCellStyle(_lastMoveBorder, _lastMoveRow);
            _lastMoveBorder = null;
            _lastMoveRow = -1;
        }
    }

    private void RollTimer_Tick(object? sender, EventArgs e)
    {
        for (var i = 0; i < DiceCount; i++)
        {
            if (!_fixedDice[i])
                _dice[i] = _random.Next(1, 7);
        }

        BuildDice();
    }

    private void MoveSelection(int delta)
    {
        if (!_selectedDice.Any(selected => selected))
        {
            _selectedDice[0] = true;
            BuildDice();
            return;
        }

        var index = Array.FindIndex(_selectedDice, selected => selected);
        _selectedDice[index] = false;
        _selectedDice[(index + delta + DiceCount) % DiceCount] = true;
        BuildDice();
    }

    private static string GetCombinationHelp(int row)
    {
        return row switch
        {
            < 6 => $"Школа {row + 1}: три кости дают 0, меньше трех - минус, больше трех - плюс.",
            6 => "Пара: максимальная пара, сумма двух одинаковых костей.",
            7 => "Пары: две разные пары.",
            8 => "Три: три одинаковые кости.",
            9 => "Фул: три плюс пара; пять одинаковых тоже считаются фулом.",
            10 => "Малая строка: 1-2-3-4-5.",
            11 => "Большая строка: 2-3-4-5-6.",
            12 => "Каре: четыре одинаковые кости плюс 20.",
            13 => "Абак: пять одинаковых костей плюс 50.",
            14 => "Сумма всех костей.",
            _ => ""
        };
    }

    private void RollButton_Click(object sender, RoutedEventArgs e) => ToggleRoll();

    private void CombinationButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: int row })
            ScoreCombination(row);
    }

    private void NewGame_Click(object sender, RoutedEventArgs e) => StartNewGame();

    private void Exit_Click(object sender, RoutedEventArgs e) => Close();

    private void Rules_Click(object sender, RoutedEventArgs e)
    {
        const string rules =
            "ABACA: правила игры\n\n" +
            "Цель: набрать больше очков, чем соперник. Игроки ходят по очереди и записывают результат в первую свободную ячейку выбранной строки.\n\n" +
            "Ход: у игрока до трех бросков. Нажмите СТАРТ/СТОП или Space. После остановки можно кликом фиксировать нужные кубики; зафиксированные кубики не перебрасываются.\n\n" +
            "Школа 1-6: считается количество выбранного номинала. Три одинаковых дают 0; меньше трех записывает минус, больше трех - плюс. Отрицательная школа влияет на общий счет по старым правилам Abaca.\n\n" +
            "Комбинации: пара, две пары, тройка, фул, малая строка 1-2-3-4-5, большая строка 2-3-4-5-6, каре, абак и сумма. Комбинация, собранная с первого броска, удваивается.\n\n" +
            "Призы: заполненная строка дает приз. В школе приз равен номиналу строки x 5, в остальных строках - максимуму строки. Заполненная колонка дает максимум колонки. Приз соперника в той же строке или колонке вычеркивается (за исключением последней колонки).\n\n" +
            "Горячие клавиши:\n" +
            "  Space        - старт / стоп броска\n" +
            "  Left / Right - выбрать кубик слева / справа\n" +
            "  Up           - зафиксировать выбранный кубик\n" +
            "  Down         - снять фиксацию выбранного кубика\n\n" +
            "Комбинации по клавишам:\n" +
            "  1 - школа 1\n" +
            "  2 - школа 2\n" +
            "  3 - школа 3\n" +
            "  4 - школа 4\n" +
            "  5 - школа 5\n" +
            "  6 - школа 6\n" +
            "  P - пара\n" +
            "  D - пары\n" +
            "  T - тройка\n" +
            "  F - фул\n" +
            "  E - малая строка\n" +
            "  S - большая строка\n" +
            "  C - каре\n" +
            "  A - абак\n" +
            "  + - сумма";
        ShowLargeMessage("Help - Rules", rules, 22);
    }

    private void About_Click(object sender, RoutedEventArgs e)
    {
        ShowLargeMessage("About", "Abaca\nНовая WPF-версия старой WinForms-игры с обновленной графикой и встроенной справкой.", 24);
    }

    private void ShowLargeMessage(string title, string message, double fontSize = 24)
    {
        var text = new TextBlock
        {
            Text = message,
            FontSize = fontSize,
            Foreground = new SolidColorBrush(Color.FromRgb(232, 240, 247)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(24)
        };
        var okButton = new Button
        {
            Content = "OK",
            Width = 120,
            Height = 46,
            FontSize = 22,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 24, 24),
            IsDefault = true
        };
        var panel = new DockPanel
        {
            Background = new SolidColorBrush(Color.FromRgb(23, 33, 43))
        };
        var scroll = new ScrollViewer
        {
            Content = text,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        };
        DockPanel.SetDock(okButton, Dock.Bottom);
        okButton.HorizontalAlignment = HorizontalAlignment.Right;
        panel.Children.Add(okButton);
        panel.Children.Add(scroll);

        var dialog = new Window
        {
            Title = title,
            Owner = this,
            Width = 720,
            Height = 430,
            MinWidth = 560,
            MinHeight = 320,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new SolidColorBrush(Color.FromRgb(23, 33, 43)),
            Content = panel
        };
        okButton.Click += (_, _) => dialog.Close();
        dialog.ShowDialog();
    }

    private void PlayDiceRollSound()
    {
        try
        {
            _diceSoundPlayer.Stop();
            if (_diceSoundPlayer.Stream is not null)
                _diceSoundPlayer.Stream.Position = 0;
            _diceSoundPlayer.Play();
        }
        catch
        {
            SystemSounds.Asterisk.Play();
        }
    }

    private static MemoryStream CreateDiceRollSound()
    {
        const int sampleRate = 22050;
        const short bitsPerSample = 16;
        const short channels = 1;
        const double durationSeconds = 0.54;
        double[] impactTimes = [0.012, 0.064, 0.132, 0.214, 0.318, 0.438];

        var sampleCount = (int)(sampleRate * durationSeconds);
        var dataSize = sampleCount * channels * bitsPerSample / 8;
        var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, System.Text.Encoding.ASCII, leaveOpen: true);
        var noise = new Random(42);

        WriteAscii(writer, "RIFF");
        writer.Write(36 + dataSize);
        WriteAscii(writer, "WAVE");
        WriteAscii(writer, "fmt ");
        writer.Write(16);
        writer.Write((short)1);
        writer.Write(channels);
        writer.Write(sampleRate);
        writer.Write(sampleRate * channels * bitsPerSample / 8);
        writer.Write((short)(channels * bitsPerSample / 8));
        writer.Write(bitsPerSample);
        WriteAscii(writer, "data");
        writer.Write(dataSize);

        for (var i = 0; i < sampleCount; i++)
        {
            var time = i / (double)sampleRate;
            var value = 0.0;

            for (var impactIndex = 0; impactIndex < impactTimes.Length; impactIndex++)
            {
                var age = time - impactTimes[impactIndex];
                if (age < 0)
                    continue;

                var strength = 1.0 - impactIndex * 0.09;
                var tapEnvelope = Math.Exp(-age * 68);
                var knockEnvelope = Math.Exp(-age * 22);
                var bodyEnvelope = Math.Exp(-age * 15);
                var grain = (noise.NextDouble() * 2 - 1) * tapEnvelope * 0.055;
                var knock = Math.Sin(2 * Math.PI * (92 + impactIndex * 5) * age) * knockEnvelope * 0.58;
                var body = Math.Sin(2 * Math.PI * (46 + impactIndex * 3) * age) * bodyEnvelope * 0.38;
                var thud = Math.Sin(2 * Math.PI * (31 + impactIndex * 2) * age) * Math.Exp(-age * 12) * 0.2;
                value += (grain + knock + body + thud) * strength;
            }

            // Gentle damping keeps the synthesized wooden knocks short and dry.
            var fadeOut = Math.Clamp((durationSeconds - time) / 0.08, 0, 1);
            value *= fadeOut;
            value = Math.Clamp(value * 0.82, -1, 1);
            writer.Write((short)(value * short.MaxValue));
        }

        stream.Position = 0;
        return stream;
    }

    private static void WriteAscii(BinaryWriter writer, string value)
    {
        foreach (var character in value)
            writer.Write((byte)character);
    }

    // Keyboard shortcuts for dice navigation and quick combination writes.
    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (_computerTurnInProgress || CurrentPlayer.IsComputer)
            return;

        if (_isRolling && e.Key != Key.Space)
            return;

        var handled = true;
        switch (e.Key)
        {
            case Key.Space:
                ToggleRoll();
                break;
            case Key.Left:
                MoveSelection(-1);
                break;
            case Key.Right:
                MoveSelection(1);
                break;
            case Key.Up:
                SetSelectedFixed(true);
                break;
            case Key.Down:
                SetSelectedFixed(false);
                break;
            case Key.D1:
                ScoreCombination(0);
                break;
            case Key.D2:
                ScoreCombination(1);
                break;
            case Key.D3:
                ScoreCombination(2);
                break;
            case Key.D4:
                ScoreCombination(3);
                break;
            case Key.D5:
                ScoreCombination(4);
                break;
            case Key.D6:
                ScoreCombination(5);
                break;
            case Key.P:
                ScoreCombination(6);
                break;
            case Key.D:
                ScoreCombination(7);
                break;
            case Key.T:
                ScoreCombination(8);
                break;
            case Key.F:
                ScoreCombination(9);
                break;
            case Key.E:
                ScoreCombination(10);
                break;
            case Key.S:
                ScoreCombination(11);
                break;
            case Key.C:
                ScoreCombination(12);
                break;
            case Key.A:
                ScoreCombination(13);
                break;
            case Key.Add:
            case Key.OemPlus:
                ScoreCombination(14);
                break;
            default:
                handled = false;
                break;
        }

        e.Handled = handled;
    }

    private void SetSelectedFixed(bool isFixed)
    {
        var index = Array.FindIndex(_selectedDice, selected => selected);
        if (index < 0)
            return;

        _fixedDice[index] = isFixed;
        BuildDice();
    }

    // Computer turn loop and simple move selection strategy.
    private async void BeginComputerTurn()
    {
        if (_computerTurnInProgress)
            return;

        _computerTurnInProgress = true;
        UpdateInputState();
        var scored = false;

        try
        {
            await Task.Delay(750);
            while (CurrentPlayer.IsComputer && _rollCount < 3)
            {
                if (_isRolling)
                {
                    StopRolling();
                    UpdateScorePanel();
                    BuildDice();
                }

                await Task.Delay(450);
                var best = GetBestComputerMove();
                if (_rollCount >= 3 || ShouldComputerStop(best))
                {
                    _computerTurnInProgress = false;
                    ScoreCombination(best.Row, true);
                    scored = true;
                    break;
                }

                ApplyComputerKeepStrategy();
                BuildDice();
                await Task.Delay(350);

                if (_fixedDice.All(fixedDie => fixedDie))
                {
                    _computerTurnInProgress = false;
                    ScoreCombination(best.Row, true);
                    scored = true;
                    break;
                }

                StartRolling();
                UpdateScorePanel();
                BuildDice();
                await Task.Delay(750);
            }

            if (!scored && CurrentPlayer.IsComputer)
            {
                if (_isRolling)
                    StopRolling();
                var best = GetBestComputerMove();
                _computerTurnInProgress = false;
                ScoreCombination(best.Row, true);
                scored = true;
            }
        }
        finally
        {
            if (!scored)
                _computerTurnInProgress = false;
            UpdateInputState();
        }
    }

    private bool ShouldComputerStop((int Row, int Score) best)
    {
        var threshold = CurrentPlayer.Difficulty switch
        {
            ComputerDifficulty.Careful => 20,
            ComputerDifficulty.Aggressive => 42,
            _ => 32
        };

        if (best.Row == 13 || best.Row == 12 && best.Score > 0)
            return true;

        return best.Score >= threshold;
    }

    private (int Row, int Score) GetBestComputerMove()
    {
        var bestRow = -1;
        var bestScore = int.MinValue;

        for (var row = 0; row < RowCount - 1; row++)
        {
            if (CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
            if (row < 6 && score < 0 && -score > CurrentPlayer.School && !CurrentPlayer.IsOnlySchoolFree())
                continue;

            var adjusted = score + GetDifficultyRowBonus(row, score);
            if (row >= 6 && score == 0)
                adjusted -= CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive ? 4 : 12;
            if (row < 6 && score < 0)
                adjusted -= CurrentPlayer.Difficulty == ComputerDifficulty.Careful ? 18 : 8;

            if (adjusted > bestScore)
            {
                bestScore = adjusted;
                bestRow = row;
            }
        }

        if (bestRow == -1)
            bestRow = Enumerable.Range(0, RowCount - 1).First(row => CurrentPlayer.GetFreeCell(row) != -1);

        return (bestRow, CalculateScore(bestRow));
    }

    private int GetDifficultyRowBonus(int row, int score)
    {
        return CurrentPlayer.Difficulty switch
        {
            ComputerDifficulty.Careful when row < 6 && score >= 0 => 7,
            ComputerDifficulty.Careful when row == 14 => 4,
            ComputerDifficulty.Aggressive when row is >= 10 and <= 13 => 12,
            ComputerDifficulty.Aggressive when row >= 6 && score > 0 => 5,
            _ => 0
        };
    }

    private void ApplyComputerKeepStrategy()
    {
        Array.Fill(_fixedDice, false);
        Array.Fill(_selectedDice, false);

        var counts = Enumerable.Range(1, 6)
            .Select(value => new { Value = value, Count = _dice.Count(die => die == value) })
            .OrderByDescending(item => item.Count)
            .ThenByDescending(item => item.Value)
            .ToArray();

        if (CurrentPlayer.Difficulty == ComputerDifficulty.Careful)
        {
            var pairValue = counts.FirstOrDefault(item => item.Count >= 2)?.Value ?? counts[0].Value;
            for (var i = 0; i < DiceCount; i++)
                _fixedDice[i] = _dice[i] == pairValue || (_rollCount >= 2 && _dice[i] >= 5);
            return;
        }

        var smallStraight = Enumerable.Range(1, 5).Count(value => _dice.Contains(value));
        var bigStraight = Enumerable.Range(2, 5).Count(value => _dice.Contains(value));
        var straightTarget = CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive ? 3 : 4;
        if (smallStraight >= straightTarget || bigStraight >= straightTarget)
        {
            var straightValues = smallStraight >= bigStraight ? Enumerable.Range(1, 5).ToHashSet() : Enumerable.Range(2, 5).ToHashSet();
            for (var i = 0; i < DiceCount; i++)
                _fixedDice[i] = straightValues.Contains(_dice[i]);
            return;
        }

        var keepValue = counts[0].Value;
        if (CurrentPlayer.Difficulty == ComputerDifficulty.Aggressive && counts[0].Count < 3)
            keepValue = counts.OrderByDescending(item => item.Value).First().Value;

        for (var i = 0; i < DiceCount; i++)
            _fixedDice[i] = _dice[i] == keepValue;
    }

    // Player state and helpers for occupied rows/columns.
    private sealed class Player
    {
        public Player(string name = "Игрок", bool isComputer = false, ComputerDifficulty difficulty = ComputerDifficulty.Normal)
        {
            Name = string.IsNullOrWhiteSpace(name) ? "Игрок" : name.Trim();
            IsComputer = isComputer;
            Difficulty = difficulty;
            for (var row = 0; row < RowCount; row++)
            for (var column = 0; column < ColumnCount; column++)
                Table[row, column] = EmptyCell;
        }

        public string Name { get; }
        public bool IsComputer { get; }
        public ComputerDifficulty Difficulty { get; }
        public int Score { get; set; }
        public int School { get; set; }
        public int[,] Table { get; } = new int[RowCount, ColumnCount];

        public bool IsOnlySchoolFree()
        {
            for (var row = 6; row < RowCount - 1; row++)
            for (var column = 0; column < ColumnCount - 1; column++)
            {
                if (Table[row, column] == EmptyCell)
                    return false;
            }

            return true;
        }

        public bool IsAllColumnBusy(int column)
        {
            for (var row = 0; row < RowCount - 1; row++)
            {
                if (Table[row, column] == EmptyCell)
                    return false;
            }

            return true;
        }

        public bool IsAllRowMainCellsBusy(int row)
        {
            for (var column = 0; column < ColumnCount - 1; column++)
            {
                if (Table[row, column] == EmptyCell)
                    return false;
            }

            return true;
        }

        public int GetFreeCell(int row)
        {
            for (var column = 0; column < ColumnCount; column++)
            {
                if (Table[row, column] == EmptyCell)
                    return column;
            }

            return -1;
        }

        public void SetBusyCell(int row, int column, int value) => Table[row, column] = value;

        public int GetMaxInRow(int row)
        {
            var max = Table[row, 0];
            for (var column = 1; column < ColumnCount - 1; column++)
            {
                if (Table[row, column] > max)
                    max = Table[row, column];
            }

            return max;
        }

        public int GetMaxInColumn(int column)
        {
            var max = Table[0, column];
            for (var row = 1; row < RowCount - 1; row++)
            {
                if (Table[row, column] > max)
                    max = Table[row, column];
            }

            return max;
        }
    }
}
