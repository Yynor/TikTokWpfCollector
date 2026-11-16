
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Playwright;

namespace TikTokWpfCollector
{
    public partial class MainWindow : Window
    {
        private const string FeedUrl = "https://www.tiktok.com/foryou";
        private const int MaxFeedSteps = 10_000_000;

        private readonly List<BloggerInfo> _bloggers = new();
        private readonly HashSet<string> _visited =
            new(StringComparer.OrdinalIgnoreCase);

        private CancellationTokenSource? _cancellationTokenSource;
        private bool _isRunning;

        public ObservableCollection<BloggerInfo> Bloggers { get; } = new();

        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;

            LoadExistingData();
            UpdateStatistics();
        }

        // =========================
        // КНОПКИ ИНТЕРФЕЙСА
        // =========================

        private async void StartButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning)
                return;

            if (!int.TryParse(TargetContactsBox.Text, out int targetContacts)
                || targetContacts <= 0)
            {
                MessageBox.Show(
                    "Укажи корректное количество контактов.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            string excelPath;

            try
            {
                excelPath = Path.GetFullPath(ExcelPathBox.Text.Trim());

                if (string.IsNullOrWhiteSpace(ExcelPathBox.Text))
                    throw new ArgumentException();
            }
            catch
            {
                MessageBox.Show(
                    "Укажи корректный путь к Excel-файлу.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);

                return;
            }

            _cancellationTokenSource = new CancellationTokenSource();

            SetRunningState(true);

            AddLog("Запуск парсера TikTok.");
            AddLog($"Цель: {targetContacts} контактов.");
            AddLog($"Excel: {excelPath}");
            AddLog($"Уже загружено профилей: {_bloggers.Count}");

            try
            {
                await RunCollectorAsync(
                    targetContacts,
                    excelPath,
                    _cancellationTokenSource.Token);
            }
            catch (OperationCanceledException)
            {
                AddLog("Работа остановлена пользователем.");
            }
            catch (Exception ex)
            {
                AddLog($"Критическая ошибка: {ex.Message}");
            }
            finally
            {
                SetRunningState(false);

                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;

                UpdateStatistics();
            }
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isRunning)
                return;

            AddLog("Отправлен запрос на остановку...");
            _cancellationTokenSource?.Cancel();
        }

        private void OpenExcelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string path = Path.GetFullPath(ExcelPathBox.Text.Trim());

                if (!File.Exists(path))
                {
                    MessageBox.Show(
                        "Excel-файл пока не создан.",
                        "Информация",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);

                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Ошибка открытия Excel",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        // =========================
        // ОСНОВНОЙ ПАРСЕР
        // =========================

        private async Task RunCollectorAsync(
            int targetContacts,
            string excelPath,
            CancellationToken cancellationToken)
        {
            string baseDir = AppContext.BaseDirectory;

            string profilePath = Path.Combine(
                baseDir,
                "tiktok_chrome_profile");

            Directory.CreateDirectory(profilePath);

            IPlaywright? playwright = null;
            IBrowserContext? context = null;

            try
            {
                AddLog("Запускаю Playwright...");

                playwright = await Playwright.CreateAsync();

                AddLog("Открываю Google Chrome...");

                context = await playwright.Chromium.LaunchPersistentContextAsync(
                    profilePath,
                    new BrowserTypeLaunchPersistentContextOptions
                    {
                        Headless = false,
                        Channel = "chrome",
                        ViewportSize = null,
                        Args = new[]
                        {
                            "--start-maximized"
                        }
                    });

                IPage page = context.Pages.Count > 0
                    ? context.Pages[0]
                    : await context.NewPageAsync();

                AddLog($"Открываю ленту: {FeedUrl}");

                await page.GotoAsync(
                    FeedUrl,
                    new PageGotoOptions
                    {
                        WaitUntil = WaitUntilState.DOMContentLoaded,
                        Timeout = 45000
                    });

                await Task.Delay(5000, cancellationToken);

                AddLog("Chrome открыт.");
                AddLog("Если требуется вход, авторизуйся вручную в Chrome.");

                int contactsFound = _bloggers.Count(HasContact);
                int step = 0;

                while (contactsFound < targetContacts
                       && step < MaxFeedSteps)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    step++;

                    AddLog($"Шаг ленты: {step}");

                    BloggerInfo? author =
                        await FindCurrentAuthorAsync(page);

                    if (author == null)
                    {
                        AddLog(
                            "Не удалось определить автора. Перехожу к следующему видео.",
                            "WARN");

                        await MoveToNextVideoAsync(
                            page,
                            cancellationToken);

                        continue;
                    }

                    string usernameKey =
                        NormalizeUsername(author.Username);

                    if (string.IsNullOrWhiteSpace(usernameKey))
                    {
                        await MoveToNextVideoAsync(
                            page,
                            cancellationToken);

                        continue;
                    }

                    if (_visited.Contains(usernameKey))
                    {
                        AddLog(
                            $"@{author.Username} уже обработан — пропуск.");

                        await MoveToNextVideoAsync(
                            page,
                            cancellationToken);

                        continue;
                    }

                    _visited.Add(usernameKey);

                    AddLog($"Найден автор: @{author.Username}", "OK");

                    IPage? profilePage = null;

                    try
                    {
                        profilePage = await context.NewPageAsync();

                        await profilePage.GotoAsync(
                            author.ProfileUrl,
                            new PageGotoOptions
                            {
                                WaitUntil = WaitUntilState.DOMContentLoaded,
                                Timeout = 30000
                            });

                        await Task.Delay(1800, cancellationToken);

                        // BIO
                        author.Bio = CleanSingleLine(
                            await ReadFirstTextAsync(
                                profilePage,
                                new[]
                                {
                                    "[data-e2e='user-bio']",
                                    "[data-e2e='user-bio-container']"
                                }));

                        // ВНЕШНЯЯ ССЫЛКА
                        author.ExternalLink =
                            await ReadProfileExternalLinkAsync(profilePage);

                        // КОЛИЧЕСТВО ПОДПИСЧИКОВ
                        author.Followers =
                            await ReadFollowersAsync(profilePage);

                        // ТЕКСТ СТРАНИЦЫ
                        string bodyText = "";

                        try
                        {
                            bodyText = await profilePage
                                .Locator("body")
                                .InnerTextAsync(
                                    new LocatorInnerTextOptions
                                    {
                                        Timeout = 5000
                                    });
                        }
                        catch (PlaywrightException ex)
                        {
                            AddLog(
                                $"Не удалось прочитать страницу: {ex.Message}",
                                "WARN");
                        }

                        string searchableText =
                            $"{author.Bio}\n" +
                            $"{author.ExternalLink}\n" +
                            $"{bodyText}";

                        // TELEGRAM И EMAIL
                        author.Telegram =
                            ExtractTelegram(searchableText);

                        author.Email =
                            ExtractEmail(searchableText);

                        AddLog($"Bio: {OrMissing(author.Bio)}");
                        AddLog($"Telegram: {OrMissing(author.Telegram)}");
                        AddLog($"Email: {OrMissing(author.Email)}");
                        AddLog($"Ссылка: {OrMissing(author.ExternalLink)}");
                        AddLog($"Подписчики: {OrMissing(author.Followers)}");

                        // СОХРАНЯЕМ КАЖДЫЙ ПРОФИЛЬ
                        _bloggers.Add(author);

                        await Dispatcher.InvokeAsync(() =>
                        {
                            Bloggers.Add(author);
                            UpdateStatistics();
                        });

                        SaveToExcel(_bloggers, excelPath);

                        if (HasContact(author))
                        {
                            contactsFound++;

                            AddLog(
                                $"Найден контакт: {contactsFound}/{targetContacts}",
                                "OK");
                        }
                        else
                        {
                            AddLog(
                                "Контакты не обнаружены. Профиль сохранён.",
                                "WARN");
                        }

                        AddLog(
                            $"Excel сохранён. Всего профилей: {_bloggers.Count}",
                            "OK");
                    }
                    catch (TimeoutException ex)
                    {
                        AddLog(
                            $"Тайм-аут профиля @{author.Username}: {ex.Message}",
                            "ERROR");
                    }
                    catch (PlaywrightException ex)
                    {
                        AddLog(
                            $"Ошибка браузера @{author.Username}: {ex.Message}",
                            "ERROR");
                    }
                    catch (Exception ex)
                    {
                        AddLog(
                            $"Ошибка профиля @{author.Username}: {ex.Message}",
                            "ERROR");
                    }
                    finally
                    {
                        if (profilePage != null)
                        {
                            try
                            {
                                await profilePage.CloseAsync();
                            }
                            catch
                            {
                                // Вкладка могла закрыться сама.
                            }
                        }

                        try
                        {
                            await page.BringToFrontAsync();
                        }
                        catch
                        {
                            // Основная вкладка могла закрыться.
                        }
                    }

                    await MoveToNextVideoAsync(
                        page,
                        cancellationToken);
                }

                SaveToExcel(_bloggers, excelPath);

                AddLog("Обработка завершена.", "OK");
                AddLog($"Проверено профилей: {_bloggers.Count}");
                AddLog(
                    $"Telegram: {_bloggers.Count(x => !string.IsNullOrWhiteSpace(x.Telegram))}");
                AddLog(
                    $"Email: {_bloggers.Count(x => !string.IsNullOrWhiteSpace(x.Email))}");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (PlaywrightException ex)
            {
                AddLog($"Ошибка Playwright/Chrome: {ex.Message}", "ERROR");
                AddLog(
                    "Проверь установку Chrome и пакет Microsoft.Playwright.",
                    "WARN");
            }
            finally
            {
                if (context != null)
                {
                    try
                    {
                        await context.CloseAsync();
                    }
                    catch
                    {
                        // Игнорируем ошибку закрытия.
                    }
                }

                playwright?.Dispose();
            }
        }

        // =========================
        // ПОИСК АВТОРА В ЛЕНТЕ
        // =========================

        private async Task<BloggerInfo?> FindCurrentAuthorAsync(IPage page)
        {
            string[] selectors =
            {
                "a[data-e2e='video-author-uniqueid']",
                "[data-e2e='video-author-uniqueid'] a",
                "a[href*='/@'][data-e2e*='author']"
            };

            foreach (string selector in selectors)
            {
                try
                {
                    ILocator locator = page.Locator(selector).First;

                    if (await locator.CountAsync() == 0 ||
                        !await locator.IsVisibleAsync())
                    {
                        continue;
                    }

                    string href =
                        (await locator.GetAttributeAsync("href") ?? "").Trim();

                    string text =
                        (await locator.InnerTextAsync()).Trim();

                    string username = NormalizeUsername(text);

                    if (string.IsNullOrWhiteSpace(username))
                        username = ExtractUsernameFromUrl(href);

                    string profileUrl =
                        NormalizeProfileUrl(href, page.Url);

                    if (!string.IsNullOrWhiteSpace(username) &&
                        !string.IsNullOrWhiteSpace(profileUrl))
                    {
                        return new BloggerInfo
                        {
                            Username = username,
                            ProfileUrl = profileUrl
                        };
                    }
                }
                catch (PlaywrightException)
                {
                    // Проверяем следующий селектор.
                }
            }

            return null;
        }

        private static string NormalizeProfileUrl(
            string href,
            string currentUrl)
        {
            if (string.IsNullOrWhiteSpace(href))
                return "";

            try
            {
                string absolute =
                    new Uri(new Uri(currentUrl), href).AbsoluteUri;

                if (!Uri.TryCreate(
                        absolute,
                        UriKind.Absolute,
                        out Uri? uri))
                {
                    return "";
                }

                if (!uri.Host.Equals(
                        "tiktok.com",
                        StringComparison.OrdinalIgnoreCase) &&
                    !uri.Host.EndsWith(
                        ".tiktok.com",
                        StringComparison.OrdinalIgnoreCase))
                {
                    return "";
                }

                Match match = Regex.Match(
                    uri.AbsolutePath,
                    @"^/@([A-Za-z0-9._-]+)(?:/|$)");

                return match.Success
                    ? $"https://www.tiktok.com/@{match.Groups[1].Value}"
                    : "";
            }
            catch
            {
                return "";
            }
        }

        private static string ExtractUsernameFromUrl(string url)
        {
            Match match = Regex.Match(
                url ?? "",
                @"/@([A-Za-z0-9._-]+)(?:/|$)");

            return match.Success
                ? match.Groups[1].Value
                : "";
        }

        private static string NormalizeUsername(string username)
        {
            return (username ?? "").Trim().TrimStart('@');
        }

        // =========================
        // ЧТЕНИЕ ДАННЫХ ПРОФИЛЯ
        // =========================

        private static async Task<string> ReadFirstTextAsync(
            IPage page,
            string[] selectors)
        {
            foreach (string selector in selectors)
            {
                try
                {
                    ILocator locator = page.Locator(selector).First;

                    if (await locator.CountAsync() > 0 &&
                        await locator.IsVisibleAsync())
                    {
                        return (await locator.InnerTextAsync()).Trim();
                    }
                }
                catch (PlaywrightException)
                {
                    // Пробуем следующий селектор.
                }
            }

            return "";
        }

        private static async Task<string> ReadFollowersAsync(IPage page)
        {
            string[] selectors =
            {
                "[data-e2e='followers-count']",
                "strong[data-e2e='followers-count']",
                "[data-e2e='user-followers']"
            };

            foreach (string selector in selectors)
            {
                try
                {
                    ILocator locator = page.Locator(selector).First;

                    if (await locator.CountAsync() > 0 &&
                        await locator.IsVisibleAsync())
                    {
                        string value =
                            (await locator.InnerTextAsync()).Trim();

                        if (!string.IsNullOrWhiteSpace(value))
                            return CleanSingleLine(value);
                    }
                }
                catch (PlaywrightException)
                {
                    // Проверяем следующий селектор.
                }
            }

            return "";
        }

        private static async Task<string> ReadProfileExternalLinkAsync(
            IPage page)
        {
            string[] selectors =
            {
                "a[data-e2e='user-link']",
                "[data-e2e='user-link'] a"
            };

            foreach (string selector in selectors)
            {
                try
                {
                    ILocator locator = page.Locator(selector);

                    int count = Math.Min(
                        await locator.CountAsync(),
                        10);

                    for (int i = 0; i < count; i++)
                    {
                        string href =
                            (await locator.Nth(i).GetAttributeAsync("href") ?? "")
                            .Trim();

                        if (!Uri.TryCreate(
                                href,
                                UriKind.Absolute,
                                out Uri? uri))
                        {
                            continue;
                        }

                        if (uri.Scheme == Uri.UriSchemeHttp ||
                            uri.Scheme == Uri.UriSchemeHttps)
                        {
                            return href;
                        }
                    }
                }
                catch (PlaywrightException)
                {
                    // Пробуем следующий селектор.
                }
            }

            return "";
        }

        // =========================
        // ИЗВЛЕЧЕНИЕ КОНТАКТОВ
        // =========================

        private static string ExtractTelegram(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            Match match = Regex.Match(
                text,
                @"(?:(?:https?://)?(?:www\.)?(?:t\.me|telegram\.me)/[A-Za-z0-9_]{5,}|(?:https?://)?(?:www\.)?telegram\.dog/[A-Za-z0-9_]{5,})",
                RegexOptions.IgnoreCase);

            if (!match.Success)
                return "";

            string result = match.Value
                .Trim()
                .TrimEnd('.', ',', ';', ')', ']', '>');

            if (!result.StartsWith(
                    "http://",
                    StringComparison.OrdinalIgnoreCase) &&
                !result.StartsWith(
                    "https://",
                    StringComparison.OrdinalIgnoreCase))
            {
                result = "https://" + result;
            }

            return result;
        }

        private static string ExtractEmail(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            Match match = Regex.Match(
                text,
                @"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
                RegexOptions.IgnoreCase);

            return match.Success
                ? match.Value.Trim().TrimEnd('.', ',', ';', ')', ']')
                : "";
        }

        private static string CleanSingleLine(string text)
        {
            return Regex.Replace(text ?? "", @"\s+", " ").Trim();
        }

        private static string OrMissing(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? "(не найдено)"
                : value;
        }

        private static bool HasContact(BloggerInfo item)
        {
            return !string.IsNullOrWhiteSpace(item.Telegram)
                || !string.IsNullOrWhiteSpace(item.Email)
                || !string.IsNullOrWhiteSpace(item.ExternalLink);
        }

        // =========================
        // ПЕРЕХОД К СЛЕДУЮЩЕМУ ВИДЕО
        // =========================

        private async Task MoveToNextVideoAsync(
            IPage page,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                await page.BringToFrontAsync();
                await page.Keyboard.PressAsync("ArrowDown");
            }
            catch (PlaywrightException ex)
            {
                AddLog(
                    $"Не удалось прокрутить ленту: {ex.Message}",
                    "WARN");
            }

            await Task.Delay(2500, cancellationToken);
        }

        // =========================
        // EXCEL
        // =========================

        private void LoadExistingData()
        {
            string path = Path.Combine(
                AppContext.BaseDirectory,
                "TikTok_Feed_Contacts.xlsx");

            if (!File.Exists(path))
                return;

            try
            {
                using var workbook = new XLWorkbook(path);

                var ws = workbook.Worksheets.FirstOrDefault();

                if (ws == null)
                    return;

                var rows = ws.RangeUsed()?.RowsUsed().Skip(1);

                if (rows == null)
                    return;

                foreach (var row in rows)
                {
                    string username = row.Cell(1).GetString().Trim();

                    if (string.IsNullOrWhiteSpace(username))
                        continue;

                    var item = new BloggerInfo
                    {
                        Username = username,
                        ProfileUrl = row.Cell(2).GetString(),
                        Telegram = row.Cell(3).GetString(),
                        Email = row.Cell(4).GetString(),
                        Bio = row.Cell(5).GetString(),
                        ExternalLink = row.Cell(6).GetString(),
                        Followers = row.Cell(7).GetString()
                    };

                    _bloggers.Add(item);
                    Bloggers.Add(item);

                    string key = NormalizeUsername(item.Username);

                    if (!string.IsNullOrWhiteSpace(key))
                        _visited.Add(key);
                }

                AddLog($"Загружено старых записей: {_bloggers.Count}");
            }
            catch (Exception ex)
            {
                AddLog($"Не удалось прочитать старый Excel: {ex.Message}", "WARN");
            }
        }

        private void SaveToExcel(
            List<BloggerInfo> list,
            string filePath)
        {
            string? directory = Path.GetDirectoryName(filePath);

            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            using var workbook = new XLWorkbook();

            var ws = workbook.Worksheets.Add("Контакты из ленты");

            string[] headers =
            {
                "Никнейм",
                "Ссылка на профиль",
                "Telegram",
                "Email",
                "Bio",
                "Внешняя ссылка",
                "Подписчики"
            };

            for (int col = 0; col < headers.Length; col++)
                ws.Cell(1, col + 1).Value = headers[col];

            var header = ws.Range(1, 1, 1, headers.Length);

            header.Style.Font.Bold = true;
            header.Style.Fill.BackgroundColor = XLColor.LightGreen;
            header.Style.Alignment.WrapText = true;

            int row = 2;

            foreach (BloggerInfo item in list)
            {
                ws.Cell(row, 1).Value = item.Username;
                ws.Cell(row, 2).Value = item.ProfileUrl;
                ws.Cell(row, 3).Value = item.Telegram;
                ws.Cell(row, 4).Value = item.Email;
                ws.Cell(row, 5).Value = item.Bio;
                ws.Cell(row, 6).Value = item.ExternalLink;
                ws.Cell(row, 7).Value = item.Followers;

                row++;
            }

            ws.SheetView.FreezeRows(1);

            ws.Column(1).Width = 24;
            ws.Column(2).Width = 38;
            ws.Column(3).Width = 34;
            ws.Column(4).Width = 30;
            ws.Column(5).Width = 45;
            ws.Column(6).Width = 38;
            ws.Column(7).Width = 16;

            ws.Column(5).Style.Alignment.WrapText = true;

            string tempPath = filePath + ".tmp.xlsx";

            workbook.SaveAs(tempPath);

            // Атомарная замена файла в пределах одного тома.
            File.Move(tempPath, filePath, true);
        }

        // =========================
        // СТАТИСТИКА И ЖУРНАЛ
        // =========================

        private void UpdateStatistics()
        {
            int telegram = _bloggers.Count(
                x => !string.IsNullOrWhiteSpace(x.Telegram));

            int email = _bloggers.Count(
                x => !string.IsNullOrWhiteSpace(x.Email));

            int contacts = _bloggers.Count(HasContact);

            ProfilesCountText.Text = _bloggers.Count.ToString();
            TelegramCountText.Text = telegram.ToString();
            EmailCountText.Text = email.ToString();
            ContactsCountText.Text = contacts.ToString();
            TableCountText.Text = $"{_bloggers.Count} записей";
        }

        private void AddLog(string message, string level = "INFO")
        {
            Dispatcher.InvokeAsync(() =>
            {
                string line =
                    $"[{DateTime.Now:HH:mm:ss}] [{level}] {message}";

                LogText.Text += line + Environment.NewLine;
            });
        }

        private void SetRunningState(bool running)
        {
            _isRunning = running;

            StartButton.IsEnabled = !running;
            StopButton.IsEnabled = running;

            TargetContactsBox.IsEnabled = !running;
            ExcelPathBox.IsEnabled = !running;
            OpenExcelButton.IsEnabled = !running;

            StatusText.Text = running
                ? "Парсер работает"
                : "Ожидание запуска";
        }
    }

    // =========================
    // МОДЕЛЬ ДАННЫХ
    // =========================

    public class BloggerInfo
    {
        public string Username { get; set; } = "";
        public string ProfileUrl { get; set; } = "";
        public string Telegram { get; set; } = "";
        public string Email { get; set; } = "";
        public string Bio { get; set; } = "";
        public string ExternalLink { get; set; } = "";
        public string Followers { get; set; } = "";
    }
}
