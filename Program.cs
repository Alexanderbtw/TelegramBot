using Newtonsoft.Json;
using Telegram.Bot;
using Telegram.Bot.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Telegram.Bot.Types.InputFiles;
using System.Drawing;
using iTextSharp.text.pdf;

using Task = System.Threading.Tasks.Task;

namespace TelegramBot
{
    internal class Program
    {
        static TelegramBotClient bot = new TelegramBotClient(System.IO.File.ReadAllText(@"D:\C#\Homework\Homework 9\TelegramBot\key.txt"));

        static string fileName = "updates.json";
        static List<BotUpdate> botUpdates = new List<BotUpdate>();

        static void Main(string[] args)
        {
            try
            {
                var botUpdatesString = System.IO.File.ReadAllText(fileName);

                botUpdates = JsonConvert.DeserializeObject<List<BotUpdate>>(botUpdatesString) ?? botUpdates;
            }
            catch(Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }

            var receiverOptions = new ReceiverOptions
            {
                AllowedUpdates = new UpdateType[]
                {
                    UpdateType.Message,
                    UpdateType.EditedMessage,

                }
            };

            bot.StartReceiving(UpdateHandler, ErrorHandler, receiverOptions);
            
            Console.ReadKey();
        }

        private static Task ErrorHandler(ITelegramBotClient arg1, Exception arg2, CancellationToken arg3)
        {
            throw new NotImplementedException();
        }

        private static async Task UpdateHandler(ITelegramBotClient bot, Update update, CancellationToken arg3)
        {
            if (update.Type == UpdateType.Message)
            {
                if (update.Message.Type == MessageType.Text)
                {
                    var _botUpdate = new BotUpdate
                    {
                        text = update.Message.Text,
                        id = update.Message.Chat.Id,
                        username = update.Message.Chat.Username
                    };

                    botUpdates.Add(_botUpdate);

                    var botUpdatesString = JsonConvert.SerializeObject(botUpdates);

                    System.IO.File.WriteAllText(fileName, botUpdatesString);

                    if (_botUpdate.text.ToLower() == "print")
                    {
                        bot.SendTextMessageAsync(_botUpdate.id, "жопа");
                        return;
                    }
                }

                else if (update.Message.Type == MessageType.Document)
                {
                    var fileId = update.Message.Document.FileId;
                    var fileInfo = await bot.GetFileAsync(fileId);
                    var filePath = fileInfo.FilePath;

                    var full_path = Path.Combine($@"D:\C#\Homework\Homework 9\TelegramBot\Data\{ update.Message.Document.FileName }");
                    await using FileStream fs = System.IO.File.OpenWrite(full_path);
                    await bot.DownloadFileAsync(filePath, fs);
                    fs.Close();

                    ConvertToPDF(full_path);

                    await Task.Delay(5000);

                    await using Stream stream = System.IO.File.OpenRead(@"D:\C#\Homework\Homework 9\TelegramBot\Data\Out.pdf");
                    await bot.SendDocumentAsync(update.Message.Chat.Id, new InputOnlineFile(stream, "Out.pdf"), "Porno");
                    return;
                }
            }
        }

        private static async void ConvertToPDF(string full_path)
        {
            var ext = Path.GetExtension(full_path);
            var out_path = @"D:\C#\Homework\Homework 9\TelegramBot\Data\Out.pdf";

            if (ext == ".docx" || ext == ".doc")
            {
                var appWord = new Microsoft.Office.Interop.Word.Application();
                if (appWord.Documents != null)
                {
                    var wordDocument = appWord.Documents.Open(full_path);
                    if (wordDocument != null)
                    {
                        wordDocument.ExportAsFixedFormat(out_path,
                        Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                        wordDocument.Close();
                    }
                    appWord.Quit();
                }
            }

            else if (ext == ".jpg")
            {
                iTextSharp.text.Rectangle pageSize = null;

                using (var srcImage = new Bitmap(full_path))
                {
                    pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
                }
                using (var ms = new MemoryStream())
                {
                    var document = new iTextSharp.text.Document(pageSize);
                    PdfWriter.GetInstance(document, ms).SetFullCompression();
                    document.Open();
                    var image = iTextSharp.text.Image.GetInstance(full_path);
                    document.Add(image);
                    document.Close();

                    System.IO.File.WriteAllBytes(out_path, ms.ToArray());
                }
            }

            System.IO.File.Delete(full_path);
        }
    }
}