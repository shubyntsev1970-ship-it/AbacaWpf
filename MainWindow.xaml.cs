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
    private const string AppVersion = "0.9.32-test";
    private const double TableLineThickness = 2.5;

    // Board geometry and table markers.
    private const int DiceCount = 5;
    private const int PlayerCount = 2;
    private const int RowCount = 16;
    private const int ColumnCount = 6;
    private const int EmptyCell = -100;
    private const double DiceSize = 100;
    private const double DiceHorizontalGap = DiceSize / 2;
    private const double DiceOutlinePadding = 7;

    // Runtime game state and UI cell references.
    private readonly Random _random = new();
    private readonly DispatcherTimer _rollTimer;
    private readonly Player[] _players = [new(), new()];
    private readonly int[] _dice = [1, 2, 3, 4, 5];
    private readonly bool[] _fixedDice = new bool[DiceCount];
    private readonly bool[] _selectedDice = new bool[DiceCount];
    private readonly TextBlock[,,] _cells = new TextBlock[PlayerCount, RowCount, ColumnCount];
    private readonly SoundPlayer _diceSoundPlayer = new(CreateDiceRollSound());
    private readonly SoundPlayer _victorySoundPlayer = new(CreateVictorySound());
    private Border? _lastMoveBorder;
    private int _lastMoveRow = -1;
    private readonly string[] _rowLabels =
    [
        "1", "2", "3", "4", "5", "6", "пара", "пары", "тройка", "фул", "м. стрит", "б. стрит", "каре", "абак", "сумма", "приз"
    ];
    private readonly string[] _rowHotkeys =
    [
        "1", "2", "3", "4", "5", "6", "P", "D", "T", "F", "E", "S", "C", "A", "+", ""
    ];

    private int _currentPlayerIndex;
    private int _rollCount;
    private bool _isRolling;
    private bool _computerTurnInProgress;
    private int _gameFlowVersion;

    // Window startup and initial control construction.
    public MainWindow()
    {
        InitializeComponent();
        Title = $"Abaca {AppVersion}";
        VersionText.Text = AppVersion;
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
        CancelCurrentGameActivity();
        var dialog = new StartGameWindow(_random, _aiTrainingLoggingEnabled) { Owner = this };
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

        SetAiTrainingLogging(dialog.AiTrainingLoggingEnabled);
        AiTrainingLogMenuItem.IsChecked = _aiTrainingLoggingEnabled;
        ResetAiTrainingGameLog();
        _currentPlayerIndex = 1;
        ClearTableValues();
        StartNextTurn();
    }

    private void CancelCurrentGameActivity()
    {
        _gameFlowVersion++;
        _computerTurnInProgress = false;
        _isRolling = false;
        _rollTimer.Stop();
        RollButton.Content = "СТАРТ";
        Array.Fill(_fixedDice, false);
        Array.Fill(_selectedDice, false);
        BuildDice();
        UpdateScorePanel();
        UpdateInputState();
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

    private void AddHeaderCell(Grid grid, string text, int row, int column)
    {
        var border = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(118, 144, 160)),
            BorderThickness = new Thickness(TableLineThickness),
            Background = row == 0 || column == 0 ? new SolidColorBrush(Color.FromRgb(33, 78, 114)) : Brushes.Transparent
        };
        var fontSize = column == 0 && row > 0 ? 22 : text.Length > 7 ? 12 : text.Length > 4 ? 14 : 16;
        if (column == 0 && row > 0 && row < RowCount)
        {
            var button = new Button
            {
                Content = CreateCaptionTextBlock(text, fontSize),
                Tag = row - 1,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(0),
                Cursor = Cursors.Hand,
                ToolTip = GetCombinationHelp(row - 1)
            };
            button.Click += CombinationButton_Click;
            border.Child = button;
        }
        else
        {
            var label = CreateCaptionTextBlock(text, fontSize);
            border.Child = label;
        }
        Grid.SetRow(border, row);
        Grid.SetColumn(border, column);
        grid.Children.Add(border);
    }

    private static TextBlock AddValueCell(Grid grid, int row, int column)
    {
        var border = new Border
        {
            BorderBrush = new SolidColorBrush(Color.FromRgb(138, 160, 174)),
            BorderThickness = new Thickness(TableLineThickness),
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

            var outlineThickness = _fixedDice[index]
                ? _selectedDice[index] ? 6 : 3
                : _selectedDice[index] ? 5 : 0;

            var outerBorder = new Border
            {
                Width = DiceSize + DiceOutlinePadding * 2,
                Height = DiceSize + DiceOutlinePadding * 2,
                Margin = new Thickness(DiceHorizontalGap / 2, _fixedDice[index] ? 0 : DiceSize, DiceHorizontalGap / 2, _fixedDice[index] ? DiceSize : 0),
                CornerRadius = new CornerRadius(14),
                Padding = new Thickness(DiceOutlinePadding),
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(outlineThickness),
                BorderBrush = _fixedDice[index]
                    ? new SolidColorBrush(Color.FromRgb(239, 68, 68))
                    : _selectedDice[index]
                        ? new SolidColorBrush(Color.FromRgb(34, 197, 94))
                        : Brushes.Transparent,
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    BlurRadius = 16,
                    ShadowDepth = 4,
                    Opacity = 0.28
                },
                RenderTransform = transform,
                RenderTransformOrigin = new Point(0.5, 0.5),
                Child = new Border
                {
                    Width = DiceSize,
                    Height = DiceSize,
                    CornerRadius = new CornerRadius(10),
                    Background = new SolidColorBrush(Color.FromRgb(16, 23, 32)),
                    ClipToBounds = true,
                    Child = CreateDiceFace(_dice[index])
                }
            };
            if (_isRolling && !_fixedDice[index])
                AnimateRollingDie(outerBorder);
            outerBorder.MouseLeftButtonDown += (_, _) => ToggleDie(index);
            DiceItems.Items.Add(outerBorder);
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
            Stretch = Stretch.UniformToFill,
            Margin = new Thickness(-7),
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
    private async void ScoreCombination(int row, bool fromComputer = false)
    {
        if (_computerTurnInProgress && !fromComputer)
            return;

        if (_isRolling)
            return;

        var column = CurrentPlayer.GetFreeCell(row);
        if (column == -1)
            return;

        var score = CalculateScore(row);
        if (fromComputer && !IsLegalComputerMove(row, score))
        {
            var fallbackRow = GetBestLegalComputerFallbackRow(row);
            if (fallbackRow >= 0)
            {
                ScoreCombination(fallbackRow, true);
                return;
            }

            return;
        }
        if (row < 6 && score < 0 && CountDice(row + 1) == 0 && HasFreeCombinationMainCell())
        {
            ShowLargeMessage("ABACA", "В школе нельзя вычеркивать строку без нужной кости, пока есть свободные клетки в комбинациях.");
            return;
        }

        if (row < 6 && score < 0 && -score > CurrentPlayer.School && !CurrentPlayer.IsOnlySchoolFree())
        {
            ShowLargeMessage("ABACA", "Недостаточно очков школы для такой записи. Выберите другую комбинацию.");
            return;
        }

        if (!fromComputer && _rollCount == 1 && row >= 6 && score == 0)
        {
            if (!ShowCrossOutConfirmation(row))
                return;
        }

        WriteTable(row, column, score);
        if (fromComputer)
        {
            await Task.Delay(320);
            await CaptureAiTrainingStepAsync("07_written_to_table", row, column, score);
        }

        if (IsGameOver())
            return;

        StartNextTurn();
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
        if (!CurrentPlayer.IsAllColumnBusy(column))
            return;

        if (CurrentPlayer.Table[RowCount - 1, column] == EmptyCell)
        {
            var value = CurrentPlayer.GetMaxInColumn(column);
            CurrentPlayer.SetBusyCell(RowCount - 1, column, value);
            CurrentPlayer.Score += value;
            WriteCell(_currentPlayerIndex, RowCount - 1, column, value);
        }

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
        border.BorderThickness = new Thickness(TableLineThickness);
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

        var isDraw = _players[0].Score == _players[1].Score;
        var winner = isDraw
            ? null
            : _players[0].Score > _players[1].Score
                ? _players[0]
                : _players[1];
        var scoreDifference = Math.Abs(_players[0].Score - _players[1].Score);
        var message = isDraw
            ? "НИЧЬЯ!"
            : $"Победил {winner!.Name}: +{scoreDifference} очков";

        UpdateScorePanel();
        ShowWinnerMessage(winner?.Name, message, scoreDifference);
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
            10 => "Малый стрит: 1-2-3-4-5.",
            11 => "Большой стрит: 2-3-4-5-6.",
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

    private void AiLogging_Click(object sender, RoutedEventArgs e)
    {
        if (sender is MenuItem item)
            SetAiTrainingLogging(item.IsChecked);
    }

    private void Exit_Click(object sender, RoutedEventArgs e) => Close();

    private void Rules_Click(object sender, RoutedEventArgs e)
    {
        const string rules =
            "ABACA: правила игры\n\n" +
            "Цель: набрать больше очков, чем соперник. Игроки ходят по очереди и записывают результат в первую свободную ячейку выбранной строки таблицы.\n\n" +
            "Ход: у игрока до трех бросков. Нажмите СТАРТ/СТОП или Space. После остановки можно кликом фиксировать нужные кубики; зафиксированные кубики не перебрасываются.\n\n" +
            "Школа 1-6: считается количество выбранного номинала. Три одинаковых дают 0; меньше трех записывает минус, больше трех - плюс. Отрицательная школа влияет на общий счет по старым правилам Abaca.\n\n" +
            "Очки в комбинациях: кроме школы, в клетку записывается сумма значений кубиков, которые составляют комбинацию. Для каре дополнительно добавляется 20 очков, для абака - 50 очков.\n\n" +
            "Комбинации: пара, две пары, тройка, фул, малый стрит 1-2-3-4-5, большой стрит 2-3-4-5-6, каре, абак и сумма. Любая комбинация, кроме школы, собранная с первого броска, удваивается.\n\n" +
            "Призы: заполненная строка таблицы дает приз. В школе приз равен номиналу строки x 5, в остальных строках таблицы - максимуму строки. Заполненная колонка дает максимум колонки. Приз соперника в той же строке таблицы или колонке вычеркивается (за исключением последней колонки).\n\n" +
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
            "  E - малый стрит\n" +
            "  S - большой стрит\n" +
            "  C - каре\n" +
            "  A - абак\n" +
            "  + - сумма\n\n" +
            "Мышь: комбинацию можно выбрать не только в нижней строке, но и кликом по названию строки в таблице.";
        ShowLargeMessage("Help - Rules", rules, 22);
    }

    private void About_Click(object sender, RoutedEventArgs e)
    {
        ShowLargeMessage("О программе", "Abaca - это настольная игра в кости, где нужно собирать выгодные комбинации и грамотно заполнять таблицу.\n\nПобеждает тот, кто лучше балансирует школу, комбинации и призы по строкам и столбцам. Играйте против другого игрока или компьютера и выбирайте самый сильный ход в каждой раздаче.", 24);
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

    private bool ShowCrossOutConfirmation(int row)
    {
        var result = false;
        var combinationName = GetRowCaption(row);
        var titleBlock = new TextBlock
        {
            Text = "Комбинация не собрана",
            FontSize = 30,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(255, 235, 180)),
            Margin = new Thickness(26, 24, 26, 8),
            TextAlignment = TextAlignment.Center
        };
        var messageBlock = new TextBlock
        {
            Text = $"Вы выбрали: {combinationName}\nВычеркнуть эту клетку?",
            FontSize = 24,
            FontWeight = FontWeights.SemiBold,
            Foreground = Brushes.White,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(28, 8, 28, 16),
            TextAlignment = TextAlignment.Center
        };
        var hintBlock = new TextBlock
        {
            Text = "Нажмите \"Подтвердить\", если это осознанное вычеркивание.",
            FontSize = 18,
            Foreground = new SolidColorBrush(Color.FromRgb(178, 204, 221)),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(28, 0, 28, 22),
            TextAlignment = TextAlignment.Center
        };
        var confirmButton = new Button
        {
            Content = "Подтвердить",
            Width = 180,
            Height = 52,
            Margin = new Thickness(8),
            Style = (Style)FindResource("CommandButton")
        };
        var cancelButton = new Button
        {
            Content = "Отказаться",
            Width = 160,
            Height = 52,
            Margin = new Thickness(8),
            Style = (Style)FindResource("CommandButton")
        };
        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(18, 0, 18, 24),
            Children = { confirmButton, cancelButton }
        };
        var panel = new StackPanel
        {
            Background = new SolidColorBrush(Color.FromRgb(16, 30, 41)),
            Children = { titleBlock, messageBlock, hintBlock, buttons }
        };
        var dialog = new Window
        {
            Title = "ABACA",
            Width = 560,
            Height = 330,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Owner = this,
            ResizeMode = ResizeMode.NoResize,
            Background = new SolidColorBrush(Color.FromRgb(16, 30, 41)),
            Content = panel
        };

        confirmButton.Click += (_, _) =>
        {
            result = true;
            dialog.Close();
        };
        cancelButton.Click += (_, _) => dialog.Close();
        dialog.ShowDialog();
        return result;
    }

    private void ShowWinnerMessage(string? winnerName, string message, int scoreDifference)
    {
        var titleText = winnerName is null ? "Ничья!" : "Победитель";
        var nameText = winnerName ?? "Победила дружба";
        var detailsText = winnerName is null
            ? "Счет равный. Отличная партия!"
            : $"+{scoreDifference} очков";

        var trophy = new Image
        {
            Source = new BitmapImage(new Uri("pack://application:,,,/Assets/WinnerTrophy.png")),
            Stretch = Stretch.UniformToFill,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };

        var trophyFrame = new Border
        {
            Height = 300,
            Margin = new Thickness(28, 24, 28, 12),
            CornerRadius = new CornerRadius(8),
            ClipToBounds = true,
            BorderThickness = new Thickness(2),
            BorderBrush = new SolidColorBrush(Color.FromRgb(255, 226, 124)),
            Child = trophy
        };

        var titleBlock = new TextBlock
        {
            Text = titleText,
            FontSize = 30,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(255, 236, 152)),
            HorizontalAlignment = HorizontalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(24, 0, 24, 4)
        };

        var winnerBlock = new TextBlock
        {
            Text = nameText,
            FontSize = 42,
            FontWeight = FontWeights.ExtraBold,
            Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
            HorizontalAlignment = HorizontalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(36, 0, 36, 2)
        };

        var detailBlock = new TextBlock
        {
            Text = detailsText,
            FontSize = 24,
            FontWeight = FontWeights.Bold,
            Foreground = new SolidColorBrush(Color.FromRgb(255, 213, 108)),
            HorizontalAlignment = HorizontalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            Margin = new Thickness(24, 0, 24, 14)
        };

        var messageBlock = new TextBlock
        {
            Text = message,
            FontSize = 20,
            Foreground = new SolidColorBrush(Color.FromRgb(232, 240, 247)),
            HorizontalAlignment = HorizontalAlignment.Center,
            TextAlignment = TextAlignment.Center,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(28, 0, 28, 18)
        };

        var okButton = new Button
        {
            Content = "OK",
            Width = 140,
            Height = 48,
            FontSize = 22,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(0, 0, 0, 24),
            IsDefault = true
        };

        var content = new StackPanel
        {
            Children =
            {
                trophyFrame,
                titleBlock,
                winnerBlock,
                detailBlock,
                messageBlock,
                okButton
            }
        };

        var decorations = new Canvas
        {
            IsHitTestVisible = false
        };
        AddCoin(decorations, 40, 452, 44, -16);
        AddCoin(decorations, 87, 494, 34, 11);
        AddCoin(decorations, 598, 446, 42, 14);
        AddCoin(decorations, 548, 506, 32, -12);
        AddSparkle(decorations, 66, 72, 22, 0);
        AddSparkle(decorations, 598, 86, 18, 0.18);
        AddSparkle(decorations, 116, 384, 16, 0.33);
        AddSparkle(decorations, 562, 382, 20, 0.48);

        var framedContent = new Border
        {
            Margin = new Thickness(20),
            CornerRadius = new CornerRadius(10),
            BorderThickness = new Thickness(4),
            BorderBrush = new LinearGradientBrush(
                Color.FromRgb(255, 241, 156),
                Color.FromRgb(184, 112, 25),
                45),
            Background = new LinearGradientBrush(
                Color.FromRgb(54, 34, 24),
                Color.FromRgb(113, 67, 20),
                90),
            Child = content
        };

        var root = new Grid
        {
            Background = new LinearGradientBrush(
                Color.FromRgb(28, 18, 16),
                Color.FromRgb(95, 55, 18),
                90),
            Children =
            {
                decorations,
                framedContent
            }
        };

        var dialog = new Window
        {
            Title = "Игра закончена",
            Owner = this,
            Width = 700,
            Height = 660,
            MinWidth = 620,
            MinHeight = 560,
            ResizeMode = ResizeMode.NoResize,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new LinearGradientBrush(
                Color.FromRgb(35, 24, 20),
                Color.FromRgb(80, 49, 18),
                90),
            Content = root
        };

        okButton.HorizontalAlignment = HorizontalAlignment.Center;
        okButton.Click += (_, _) => dialog.Close();
        PlayVictorySound();
        dialog.ShowDialog();
    }

    private static void AddCoin(Canvas canvas, double left, double top, double size, double angle)
    {
        var coin = new Ellipse
        {
            Width = size,
            Height = size * 0.42,
            Fill = new RadialGradientBrush(
                Color.FromRgb(255, 245, 159),
                Color.FromRgb(213, 136, 24)),
            Stroke = new SolidColorBrush(Color.FromRgb(255, 222, 95)),
            StrokeThickness = 2,
            RenderTransform = new RotateTransform(angle, size / 2, size * 0.21)
        };

        Canvas.SetLeft(coin, left);
        Canvas.SetTop(coin, top);
        canvas.Children.Add(coin);
    }

    private static void AddSparkle(Canvas canvas, double left, double top, double size, double delaySeconds)
    {
        var sparkle = new Polygon
        {
            Points =
            [
                new Point(size / 2, 0),
                new Point(size * 0.62, size * 0.38),
                new Point(size, size / 2),
                new Point(size * 0.62, size * 0.62),
                new Point(size / 2, size),
                new Point(size * 0.38, size * 0.62),
                new Point(0, size / 2),
                new Point(size * 0.38, size * 0.38)
            ],
            Fill = new SolidColorBrush(Color.FromRgb(255, 246, 172)),
            Opacity = 0.82,
            RenderTransformOrigin = new Point(0.5, 0.5),
            RenderTransform = new ScaleTransform(0.78, 0.78)
        };

        Canvas.SetLeft(sparkle, left);
        Canvas.SetTop(sparkle, top);
        canvas.Children.Add(sparkle);

        var opacity = new DoubleAnimation(0.38, 1, TimeSpan.FromMilliseconds(760))
        {
            AutoReverse = true,
            BeginTime = TimeSpan.FromSeconds(delaySeconds),
            RepeatBehavior = RepeatBehavior.Forever
        };
        sparkle.BeginAnimation(OpacityProperty, opacity);

        if (sparkle.RenderTransform is ScaleTransform scale)
        {
            var pulse = new DoubleAnimation(0.78, 1.2, TimeSpan.FromMilliseconds(760))
            {
                AutoReverse = true,
                BeginTime = TimeSpan.FromSeconds(delaySeconds),
                RepeatBehavior = RepeatBehavior.Forever
            };
            scale.BeginAnimation(ScaleTransform.ScaleXProperty, pulse);
            scale.BeginAnimation(ScaleTransform.ScaleYProperty, pulse);
        }
    }

    private void PlayVictorySound()
    {
        try
        {
            _victorySoundPlayer.Stop();
            if (_victorySoundPlayer.Stream is not null)
                _victorySoundPlayer.Stream.Position = 0;
            _victorySoundPlayer.Play();
        }
        catch
        {
            SystemSounds.Exclamation.Play();
        }
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

    private static MemoryStream CreateVictorySound()
    {
        const int sampleRate = 22050;
        const short bitsPerSample = 16;
        const short channels = 1;
        const double durationSeconds = 1.15;
        double[] noteStarts = [0.02, 0.2, 0.38, 0.64];
        double[] noteFrequencies = [523.25, 659.25, 783.99, 1046.5];

        var sampleCount = (int)(sampleRate * durationSeconds);
        var dataSize = sampleCount * channels * bitsPerSample / 8;
        var stream = new MemoryStream();
        using var writer = new BinaryWriter(stream, System.Text.Encoding.ASCII, leaveOpen: true);

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

            for (var noteIndex = 0; noteIndex < noteStarts.Length; noteIndex++)
            {
                var age = time - noteStarts[noteIndex];
                if (age < 0)
                    continue;

                var envelope = Math.Min(age / 0.035, 1) * Math.Exp(-age * (noteIndex == noteStarts.Length - 1 ? 2.4 : 5.4));
                var frequency = noteFrequencies[noteIndex];
                var bell = Math.Sin(2 * Math.PI * frequency * age) * 0.42;
                var bright = Math.Sin(2 * Math.PI * frequency * 2 * age) * 0.13;
                var warm = Math.Sin(2 * Math.PI * frequency * 0.5 * age) * 0.1;
                value += (bell + bright + warm) * envelope;
            }

            var fadeOut = Math.Clamp((durationSeconds - time) / 0.16, 0, 1);
            value = Math.Clamp(value * fadeOut * 0.82, -1, 1);
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

        var key = GetShortcutKey(e);
        if (_isRolling && key != Key.Space)
            return;

        var handled = true;
        switch (key)
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

    private static Key GetShortcutKey(KeyEventArgs e)
    {
        if (e.Key == Key.System)
            return e.SystemKey;

        if (e.Key == Key.ImeProcessed)
            return e.ImeProcessedKey;

        return e.Key;
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

        var flowVersion = _gameFlowVersion;
        _computerTurnInProgress = true;
        UpdateInputState();
        var scored = false;
        BeginAiTrainingTurnLog();

        try
        {
            await CaptureAiTrainingStepAsync("01_before_roll_1");
            if (flowVersion != _gameFlowVersion)
                return;
            await Task.Delay(750);
            if (flowVersion != _gameFlowVersion)
                return;
            while (CurrentPlayer.IsComputer && _rollCount < 3)
            {
                if (flowVersion != _gameFlowVersion)
                    return;
                if (_isRolling)
                {
                    StopRolling();
                    UpdateScorePanel();
                    BuildDice();
                    await CaptureAiTrainingStepAsync(GetAiTrainingAfterRollStepName());
                    if (flowVersion != _gameFlowVersion)
                        return;
                }

                await Task.Delay(450);
                if (flowVersion != _gameFlowVersion)
                    return;
                var best = GetBestComputerMove();
                if (_rollCount >= 3 || ShouldComputerStop(best))
                {
                    SetAiTrainingDecision(best.Row, CalculateScore(best.Row), "stop");
                    _computerTurnInProgress = false;
                    ScoreCombination(best.Row, true);
                    scored = true;
                    break;
                }

                ApplyComputerKeepStrategy();
                BuildDice();
                await CaptureAiTrainingStepAsync(GetAiTrainingKeepStepName());
                if (flowVersion != _gameFlowVersion)
                    return;
                await Task.Delay(350);
                if (flowVersion != _gameFlowVersion)
                    return;

                if (_fixedDice.All(fixedDie => fixedDie))
                {
                    SetAiTrainingDecision(best.Row, CalculateScore(best.Row), "all_dice_fixed");
                    _computerTurnInProgress = false;
                    ScoreCombination(best.Row, true);
                    scored = true;
                    break;
                }

                StartRolling();
                UpdateScorePanel();
                BuildDice();
                await Task.Delay(750);
                if (flowVersion != _gameFlowVersion)
                    return;
            }

            if (!scored && CurrentPlayer.IsComputer)
            {
                if (_isRolling)
                    StopRolling();
                UpdateScorePanel();
                BuildDice();
                await CaptureAiTrainingStepAsync(GetAiTrainingAfterRollStepName());
                if (flowVersion != _gameFlowVersion)
                    return;
                var best = GetBestComputerMove();
                SetAiTrainingDecision(best.Row, CalculateScore(best.Row), "fallback");
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

    private int GetBestLegalComputerFallbackRow(int excludedRow)
    {
        for (var row = 0; row < RowCount - 1; row++)
        {
            if (row == excludedRow || CurrentPlayer.GetFreeCell(row) == -1)
                continue;

            var score = CalculateScore(row);
            if (IsLegalComputerMove(row, score))
                return row;
        }

        return -1;
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
            for (var column = 0; column < ColumnCount - 1; column++)
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
