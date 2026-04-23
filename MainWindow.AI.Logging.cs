using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace AbacaWpf;

public partial class MainWindow
{
    private const string AiTrainingLogFolderName = "AI_TrainingLogs";
    private const string AiTrainingProjectRoot = @"D:\Project\AbacaWpf";

    private bool _aiTrainingLoggingEnabled;
    private string? _aiTrainingGameFolder;
    private string? _aiTrainingTurnFolder;
    private int _aiTrainingTurnNumber;
    private readonly List<string> _aiTrainingTurnLines = [];

    private void SetAiTrainingLogging(bool isEnabled)
    {
        _aiTrainingLoggingEnabled = isEnabled;
        if (!isEnabled)
        {
            _aiTrainingTurnFolder = null;
            return;
        }

        EnsureAiTrainingGameFolder();
    }

    private void ResetAiTrainingGameLog()
    {
        _aiTrainingTurnFolder = null;
        _aiTrainingTurnNumber = 0;
        _aiTrainingTurnLines.Clear();
        if (_aiTrainingLoggingEnabled)
            _aiTrainingGameFolder = CreateAiTrainingGameFolder();
        else
            _aiTrainingGameFolder = null;
    }

    private void BeginAiTrainingTurnLog()
    {
        if (!_aiTrainingLoggingEnabled || !CurrentPlayer.IsComputer)
            return;

        var gameFolder = EnsureAiTrainingGameFolder();
        _aiTrainingTurnNumber++;
        _aiTrainingTurnFolder = Path.Combine(gameFolder, $"Turn_{_aiTrainingTurnNumber:000}");
        Directory.CreateDirectory(_aiTrainingTurnFolder);
        _aiTrainingTurnLines.Clear();
        AddAiTrainingInfoLine($"Version: {AppVersion}");
        AddAiTrainingInfoLine($"Turn: {_aiTrainingTurnNumber:000}");
        AddAiTrainingInfoLine($"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        AddAiTrainingInfoLine($"Computer: {CurrentPlayer.Name}");
        AddAiTrainingInfoLine($"Opponent: {_players[1 - _currentPlayerIndex].Name}");
        AddAiTrainingInfoLine("");
    }

    private void SetAiTrainingDecision(int row, int score, string reason)
    {
        if (!_aiTrainingLoggingEnabled || _aiTrainingTurnFolder is null)
            return;

        AddAiTrainingInfoLine($"Decision: {GetRowCaption(row)}; score={score}; reason={reason}; roll={_rollCount}; dice={FormatDiceValues()}");
    }

    private async Task CaptureAiTrainingStepAsync(string stage, int? row = null, int? column = null, int? score = null)
    {
        if (!_aiTrainingLoggingEnabled || _aiTrainingTurnFolder is null)
            return;

        await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);

        UpdateLayout();
        var width = Math.Max(1, (int)Math.Ceiling(ActualWidth));
        var height = Math.Max(1, (int)Math.Ceiling(ActualHeight));
        var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(this);

        var fileName = $"{SanitizePathPart(stage)}.png";
        var path = Path.Combine(_aiTrainingTurnFolder, fileName);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var stream = File.Create(path))
        {
            encoder.Save(stream);
        }

        var details = new StringBuilder();
        details.Append($"{DateTime.Now:HH:mm:ss.fff} {fileName}: {stage}; ");
        details.Append($"roll={_rollCount}; dice={FormatDiceValues()}; fixed={FormatBoolFlags(_fixedDice)}; selected={FormatBoolFlags(_selectedDice)}");
        if (row is not null)
            details.Append($"; row={GetRowCaption(row.Value)}");
        if (column is not null)
            details.Append($"; column={column.Value + 1}");
        if (score is not null)
            details.Append($"; score={score.Value}");
        AddAiTrainingInfoLine(details.ToString());
    }

    private string GetAiTrainingAfterRollStepName()
    {
        return _rollCount switch
        {
            1 => "02_after_roll_1",
            2 => "04_after_roll_2",
            3 => "06_after_roll_3",
            _ => $"after_roll_{_rollCount}"
        };
    }

    private string GetAiTrainingKeepStepName()
    {
        return _rollCount switch
        {
            1 => "03_keep_before_roll_2",
            2 => "05_keep_before_roll_3",
            _ => $"keep_after_roll_{_rollCount}"
        };
    }

    private string EnsureAiTrainingGameFolder()
    {
        _aiTrainingGameFolder ??= CreateAiTrainingGameFolder();
        return _aiTrainingGameFolder;
    }

    private static string CreateAiTrainingGameFolder()
    {
        var root = Directory.Exists(AiTrainingProjectRoot)
            ? Path.Combine(AiTrainingProjectRoot, AiTrainingLogFolderName)
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Abaca", AiTrainingLogFolderName);

        Directory.CreateDirectory(root);
        var folderName = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss", CultureInfo.InvariantCulture);
        var folder = Path.Combine(root, folderName);
        Directory.CreateDirectory(folder);
        return folder;
    }

    private void AddAiTrainingInfoLine(string line)
    {
        _aiTrainingTurnLines.Add(line);
        WriteAiTrainingInfoFile();
    }

    private void WriteAiTrainingInfoFile()
    {
        if (_aiTrainingTurnFolder is null)
            return;

        File.WriteAllLines(Path.Combine(_aiTrainingTurnFolder, "turn_info.txt"), _aiTrainingTurnLines, Encoding.UTF8);
    }

    private string FormatDiceValues()
    {
        return string.Join("-", _dice);
    }

    private static string FormatBoolFlags(bool[] values)
    {
        return string.Join("", values.Select(value => value ? "1" : "0"));
    }

    private static string SanitizePathPart(string value)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var builder = new StringBuilder(value.Length);
        foreach (var ch in value)
            builder.Append(invalid.Contains(ch) ? '_' : ch);
        return builder.ToString();
    }
}
