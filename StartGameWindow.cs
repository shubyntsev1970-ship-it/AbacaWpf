using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace AbacaWpf;

public sealed class StartGameWindow : Window
{
    // Состояние стартовых бросков: пять маленьких кубиков и суммы двух игроков.
    private readonly Random _random;
    private readonly DispatcherTimer _rollTimer;
    private readonly int[] _rollingDice = new int[5];
    private readonly TextBlock[] _firstDiceTexts = new TextBlock[5];
    private readonly TextBlock[] _secondDiceTexts = new TextBlock[5];
    private readonly TextBox _firstNameBox = new() { MinWidth = 210, Height = 32, FontSize = 15 };
    private readonly TextBox _secondNameBox = new() { MinWidth = 210, Height = 32, FontSize = 15 };
    private readonly CheckBox _computerModeBox = new()
    {
        Content = "Играть против компьютера",
        Foreground = Brushes.White,
        FontSize = 15,
        Margin = new Thickness(0, 10, 0, 0)
    };
    private readonly ComboBox _difficultyBox = new()
    {
        Width = 145,
        Height = 28,
        IsEnabled = false,
        Margin = new Thickness(12, 8, 0, 0)
    };
    private readonly CheckBox _aiTrainingLogBox = new()
    {
        Content = "AI training log",
        Foreground = Brushes.White,
        FontSize = 15,
        IsEnabled = false,
        Margin = new Thickness(0, 6, 0, 0)
    };
    private readonly TextBlock _firstRollText = new() { Text = "0", FontSize = 36, FontWeight = FontWeights.Bold };
    private readonly TextBlock _secondRollText = new() { Text = "0", FontSize = 36, FontWeight = FontWeights.Bold };
    private readonly TextBlock _rollStatusText = new()
    {
        Foreground = new SolidColorBrush(Color.FromRgb(232, 240, 247)),
        FontSize = 17,
        FontWeight = FontWeights.SemiBold,
        TextAlignment = TextAlignment.Center,
        TextWrapping = TextWrapping.Wrap,
        Margin = new Thickness(0, 4, 0, 0)
    };
    private readonly Button _rollButton = new()
    {
        Content = "Бросок игрока 1",
        Width = 170,
        Height = 40,
        FontSize = 15,
        Margin = new Thickness(0, 0, 10, 0)
    };
    private readonly Button _startButton = new()
    {
        Content = "Начать",
        Width = 118,
        Height = 40,
        FontSize = 15,
        IsDefault = true,
        IsEnabled = false
    };
    private int _firstRoll;
    private int _secondRoll;
    private int _nextRollPlayer;
    private bool _isRolling;
    private bool _mustReroll;
    private bool _computerStartRollInProgress;

    // Окно собирается кодом, потому что элементы меняют состояние во время стартовых бросков.
    public StartGameWindow(Random random, bool aiTrainingLoggingEnabled = false)
    {
        _random = random;
        _rollTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(90) };
        _rollTimer.Tick += (_, _) => UpdateRollingDice();

        Title = "ABACA - новый матч";
        Width = 760;
        Height = 610;
        ResizeMode = ResizeMode.NoResize;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        Background = new SolidColorBrush(Color.FromRgb(16, 24, 32));
        _aiTrainingLogBox.IsChecked = aiTrainingLoggingEnabled;
        Content = BuildContent();
        _firstNameBox.TextChanged += (_, _) => UpdateRollButtonText();
        _secondNameBox.TextChanged += (_, _) => UpdateRollButtonText();
        _firstNameBox.KeyDown += NameBox_KeyDown;
        _secondNameBox.KeyDown += NameBox_KeyDown;
        _difficultyBox.SelectionChanged += (_, _) =>
        {
            _firstNameBox.Focus();
            _firstNameBox.SelectAll();
        };
        Loaded += (_, _) =>
        {
            _firstNameBox.Focus();
            _firstNameBox.SelectAll();
        };
        ResetStartRolls();
    }

    // Эти свойства читает MainWindow после закрытия стартового окна.
    public string FirstPlayerName => _firstNameBox.Text.Trim();
    public string SecondPlayerName => PlayAgainstComputer ? "Компьютер" : _secondNameBox.Text.Trim();
    public bool SecondPlayerStarts => _secondRoll > _firstRoll;
    public bool PlayAgainstComputer => _computerModeBox.IsChecked == true;
    public bool AiTrainingLoggingEnabled => PlayAgainstComputer && _aiTrainingLogBox.IsChecked == true;
    public ComputerDifficulty ComputerDifficulty => _difficultyBox.SelectedIndex switch
    {
        0 => ComputerDifficulty.Careful,
        2 => ComputerDifficulty.Aggressive,
        _ => ComputerDifficulty.Normal
    };

    // Строит форму выбора игроков, сложности компьютера и видимых стартовых бросков.
    private Grid BuildContent()
    {
        var root = new Grid { Margin = new Thickness(26) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var title = new TextBlock
        {
            Text = "Кто ходит первым?",
            Foreground = new SolidColorBrush(Color.FromRgb(246, 211, 101)),
            FontSize = 28,
            FontWeight = FontWeights.Black,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 0, 0, 18)
        };
        root.Children.Add(title);

        var players = new UniformGrid { Columns = 2, Margin = new Thickness(0, 8, 0, 18) };
        players.Children.Add(BuildPlayerBox("Первый игрок", _firstNameBox, _firstRollText, _firstDiceTexts));
        players.Children.Add(BuildPlayerBox("Второй игрок", _secondNameBox, _secondRollText, _secondDiceTexts));
        Grid.SetRow(players, 1);
        root.Children.Add(players);

        Grid.SetRow(_rollStatusText, 2);
        root.Children.Add(_rollStatusText);

        _difficultyBox.Items.Add("Осторожный");
        _difficultyBox.Items.Add("Обычный");
        _difficultyBox.Items.Add("Агрессивный");
        _difficultyBox.SelectedIndex = 1;

        _computerModeBox.Checked += (_, _) =>
        {
            _secondNameBox.Text = "Компьютер";
            _secondNameBox.IsEnabled = false;
            _difficultyBox.IsEnabled = true;
            _aiTrainingLogBox.IsEnabled = true;
            UpdateRollButtonText();
            _firstNameBox.Focus();
            _firstNameBox.SelectAll();
        };
        _computerModeBox.Unchecked += (_, _) =>
        {
            _secondNameBox.IsEnabled = true;
            _secondNameBox.Text = "";
            _difficultyBox.IsEnabled = false;
            _aiTrainingLogBox.IsEnabled = false;
            _aiTrainingLogBox.IsChecked = false;
            UpdateRollButtonText();
        };
        var optionsPanel = new StackPanel { Orientation = Orientation.Vertical, VerticalAlignment = VerticalAlignment.Center };
        var computerOptionsPanel = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
        computerOptionsPanel.Children.Add(_computerModeBox);
        computerOptionsPanel.Children.Add(_difficultyBox);
        optionsPanel.Children.Add(computerOptionsPanel);
        optionsPanel.Children.Add(_aiTrainingLogBox);
        Grid.SetRow(optionsPanel, 3);
        root.Children.Add(optionsPanel);

        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 18, 0, 0) };
        _rollButton.Click += (_, _) => ToggleVisibleRoll();
        _startButton.Click += (_, _) =>
        {
            if (!HasRequiredNames())
            {
                return;
            }

            DialogResult = true;
            Close();
        };
        buttons.Children.Add(_rollButton);
        buttons.Children.Add(_startButton);
        Grid.SetRow(buttons, 4);
        root.Children.Add(buttons);

        return root;
    }

    private static Border BuildPlayerBox(string title, TextBox nameBox, TextBlock rollText, TextBlock[] diceTexts)
    {
        var panel = new StackPanel();
        panel.Children.Add(new TextBlock
        {
            Text = title,
            Foreground = new SolidColorBrush(Color.FromRgb(187, 208, 221)),
            FontWeight = FontWeights.SemiBold,
            FontSize = 15
        });
        panel.Children.Add(nameBox);
        panel.Children.Add(new TextBlock
        {
            Text = "Пять стартовых кубиков",
            Foreground = new SolidColorBrush(Color.FromRgb(143, 179, 201)),
            FontSize = 15,
            Margin = new Thickness(0, 16, 0, 6)
        });
        panel.Children.Add(BuildDiceStrip(diceTexts));
        panel.Children.Add(new TextBlock
        {
            Text = "Сумма",
            Foreground = new SolidColorBrush(Color.FromRgb(143, 179, 201)),
            FontSize = 15,
            Margin = new Thickness(0, 14, 0, 0),
            HorizontalAlignment = HorizontalAlignment.Center
        });
        rollText.Foreground = Brushes.White;
        rollText.HorizontalAlignment = HorizontalAlignment.Center;
        rollText.Margin = new Thickness(0, 2, 0, 0);
        panel.Children.Add(rollText);

        return new Border
        {
            Margin = new Thickness(8),
            Padding = new Thickness(18),
            MinHeight = 250,
            CornerRadius = new CornerRadius(8),
            Background = new SolidColorBrush(Color.FromRgb(23, 33, 43)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(49, 80, 109)),
            BorderThickness = new Thickness(1),
            Child = panel
        };
    }

    private static UniformGrid BuildDiceStrip(TextBlock[] diceTexts)
    {
        var diceStrip = new UniformGrid { Columns = 5, Height = 42, Margin = new Thickness(0, 0, 0, 2) };
        for (var i = 0; i < diceTexts.Length; i++)
        {
            var dieText = new TextBlock
            {
                Text = "0",
                Foreground = new SolidColorBrush(Color.FromRgb(18, 26, 34)),
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            diceTexts[i] = dieText;
            diceStrip.Children.Add(new Border
            {
                Margin = new Thickness(3),
                CornerRadius = new CornerRadius(6),
                Background = new LinearGradientBrush(Color.FromRgb(255, 255, 255), Color.FromRgb(219, 229, 237), 45),
                BorderBrush = new SolidColorBrush(Color.FromRgb(187, 208, 221)),
                BorderThickness = new Thickness(1),
                Child = dieText
            });
        }

        return diceStrip;
    }

    // Возвращает стартовые броски в исходное состояние 0:0 и снова разрешает ввод имен.
    private void ResetStartRolls()
    {
        _rollTimer.Stop();
        _isRolling = false;
        _mustReroll = false;
        _nextRollPlayer = 0;
        _firstRoll = 0;
        _secondRoll = 0;
        _firstRollText.Text = "0";
        _secondRollText.Text = "0";
        SetDiceValues(_firstDiceTexts, [0, 0, 0, 0, 0]);
        SetDiceValues(_secondDiceTexts, [0, 0, 0, 0, 0]);
        SetNameInputsEnabled(true);
        _startButton.IsEnabled = false;
        _rollStatusText.Foreground = new SolidColorBrush(Color.FromRgb(232, 240, 247));
        _rollStatusText.Text = HasRequiredNames()
            ? "Нажмите кнопку броска, затем остановите кубики."
            : "Введите имена игроков перед стартовыми бросками.";
        UpdateRollButtonText();
    }

    // Запускает или останавливает видимый бросок текущего стартового игрока.
    private void ToggleVisibleRoll()
    {
        if (_computerStartRollInProgress)
            return;

        if (!HasRequiredNames())
        {
            _rollStatusText.Foreground = new SolidColorBrush(Color.FromRgb(246, 211, 101));
            _rollStatusText.Text = "Сначала введите имена игроков.";
            UpdateRollButtonText();
            return;
        }

        if (_isRolling)
        {
            StopVisibleRoll();
            return;
        }

        if (_mustReroll)
        {
            ResetStartRolls();
        }

        StartVisibleRoll();
    }

    private void StartVisibleRoll()
    {
        SetNameInputsEnabled(false);
        _isRolling = true;
        _startButton.IsEnabled = false;
        _rollButton.Content = "Стоп";
        _rollStatusText.Foreground = new SolidColorBrush(Color.FromRgb(232, 240, 247));
        _rollStatusText.Text = $"Бросает: {GetRollingPlayerName()}. Нажмите «Стоп», чтобы зафиксировать сумму.";
        UpdateRollingDice();
        _rollTimer.Start();
    }

    // Фиксирует сумму остановленного стартового броска и передает ход второму игроку.
    private void StopVisibleRoll()
    {
        _rollTimer.Stop();
        _isRolling = false;
        var sum = _rollingDice.Sum();

        if (_nextRollPlayer == 0)
        {
            _firstRoll = sum;
            _firstRollText.Text = sum.ToString();
            SetDiceValues(_firstDiceTexts, _rollingDice);
            _nextRollPlayer = 1;
            _rollStatusText.Text = $"Сумма {FirstPlayerName}: {sum}. Теперь бросает {SecondPlayerName}.";
            UpdateRollButtonText();
            if (PlayAgainstComputer)
            {
                BeginComputerStartRoll();
            }

            return;
        }

        _secondRoll = sum;
        _secondRollText.Text = sum.ToString();
        SetDiceValues(_secondDiceTexts, _rollingDice);
        EvaluateStartRolls();
    }

    // После двух стартовых бросков решает: можно начинать игру или нужен переброс при равенстве.
    private void EvaluateStartRolls()
    {
        if (_firstRoll == _secondRoll)
        {
            _mustReroll = true;
            _nextRollPlayer = 0;
            _startButton.IsEnabled = false;
            _rollButton.Content = "Перебросить";
            _rollButton.IsEnabled = true;
            _rollStatusText.Foreground = new SolidColorBrush(Color.FromRgb(246, 211, 101));
            _rollStatusText.Text = $"Равенство: {_firstRoll} : {_secondRoll}. Нужно перебросить стартовые кубики.";
            return;
        }

        _mustReroll = false;
        _startButton.IsEnabled = true;
        _rollButton.Content = "Перебросить";
        _rollButton.IsEnabled = false;
        _rollStatusText.Foreground = new SolidColorBrush(Color.FromRgb(232, 240, 247));
        var starterName = _secondRoll > _firstRoll ? SecondPlayerName : FirstPlayerName;
        _rollStatusText.Text = $"Броски: {_firstRoll} : {_secondRoll}. Первым ходит: {starterName}.";
    }

    // В режиме против компьютера второй стартовый бросок выполняется автоматически с небольшой паузой.
    private async void BeginComputerStartRoll()
    {
        _computerStartRollInProgress = true;
        _rollButton.IsEnabled = false;
        _rollStatusText.Text = "Компьютер готовит стартовый бросок...";

        await Task.Delay(500);
        if (_nextRollPlayer != 1)
        {
            _computerStartRollInProgress = false;
            UpdateRollButtonText();
            return;
        }

        StartVisibleRoll();
        await Task.Delay(1100);
        if (_isRolling && _nextRollPlayer == 1)
        {
            StopVisibleRoll();
        }

        _computerStartRollInProgress = false;
        UpdateRollButtonText();
    }

    // Синхронизирует доступность кнопок и полей с текущим этапом стартового окна.
    private void SetNameInputsEnabled(bool isEnabled)
    {
        _firstNameBox.IsEnabled = isEnabled;
        _secondNameBox.IsEnabled = isEnabled && !PlayAgainstComputer;
        _computerModeBox.IsEnabled = isEnabled;
        _difficultyBox.IsEnabled = isEnabled && PlayAgainstComputer;
        _aiTrainingLogBox.IsEnabled = isEnabled && PlayAgainstComputer;
    }

    private void UpdateRollingDice()
    {
        for (var i = 0; i < _rollingDice.Length; i++)
        {
            _rollingDice[i] = _random.Next(1, 7);
        }

        SetDiceValues(_nextRollPlayer == 0 ? _firstDiceTexts : _secondDiceTexts, _rollingDice);
    }

    private void UpdateRollButtonText()
    {
        if (_isRolling)
        {
            _rollButton.Content = "Стоп";
            _rollButton.IsEnabled = true;
            return;
        }

        if (_firstRoll > 0 && _secondRoll > 0 && !_mustReroll)
        {
            _rollButton.Content = "Перебросить";
            _rollButton.IsEnabled = false;
            return;
        }

        if (!HasRequiredNames())
        {
            _rollButton.Content = "Введите имена";
            _rollButton.IsEnabled = false;
            _startButton.IsEnabled = false;
            if (_firstRoll == 0 && _secondRoll == 0)
            {
                _rollStatusText.Foreground = new SolidColorBrush(Color.FromRgb(246, 211, 101));
                _rollStatusText.Text = "Введите имена игроков перед стартовыми бросками.";
            }

            return;
        }

        _rollButton.IsEnabled = true;
        _rollButton.Content = _nextRollPlayer == 0
            ? $"Бросок {FirstPlayerName}"
            : $"Бросок {SecondPlayerName}";
    }

    private string GetRollingPlayerName() => _nextRollPlayer == 0 ? FirstPlayerName : SecondPlayerName;

    private bool HasRequiredNames() =>
        !string.IsNullOrWhiteSpace(_firstNameBox.Text)
        && (PlayAgainstComputer || !string.IsNullOrWhiteSpace(_secondNameBox.Text));

    private void NameBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;

        e.Handled = true;
        if (sender == _firstNameBox && !PlayAgainstComputer)
        {
            _secondNameBox.Focus();
            _secondNameBox.SelectAll();
            return;
        }

        if (HasRequiredNames() && _firstRoll == 0 && _secondRoll == 0)
        {
            ToggleVisibleRoll();
        }
        else
        {
            UpdateRollButtonText();
        }
    }

    private static void SetDiceValues(TextBlock[] diceTexts, IReadOnlyList<int> values)
    {
        for (var i = 0; i < diceTexts.Length; i++)
        {
            diceTexts[i].Text = values[i].ToString();
        }
    }
}
