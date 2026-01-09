using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Diagnostics;

namespace XpOllamaTerminal
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private const string DefaultEndpoint = "http://127.0.0.1:11434";
        private const string DefaultModel = "llava:7b";

        private RichTextBox termBox;
        private RichTextBox chatBox;
        private TextBox inputBox;
        private TextBox endpointBox;
        private TextBox modelBox;
        private Label statusLabel;
        private Button sendAiButton;
        private Button runCmdButton;

        private Process cmdProc;
        private StreamWriter cmdIn;
        private Thread cmdReaderThread;

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern int GetPrivateProfileString(
            string section, string key, string def,
            StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern bool WritePrivateProfileString(
            string section, string key, string val, string filePath);

        public MainForm()
        {
            Text = "XP Ollama Terminal (.NET 4.0)";
            Width = 940;
            Height = 490;
            StartPosition = FormStartPosition.CenterScreen;

            termBox = new RichTextBox();
            termBox.ReadOnly = true;
            termBox.Multiline = true;
            termBox.ScrollBars = RichTextBoxScrollBars.Vertical;
            termBox.SetBounds(10, 10, 520, 260);

            chatBox = new RichTextBox();
            chatBox.ReadOnly = true;
            chatBox.Multiline = true;
            chatBox.ScrollBars = RichTextBoxScrollBars.Vertical;
            chatBox.SetBounds(540, 10, 370, 260);

            inputBox = new TextBox();
            inputBox.Multiline = true;
            inputBox.ScrollBars = ScrollBars.Vertical;
            inputBox.SetBounds(540, 280, 370, 80);

            sendAiButton = new Button();
            sendAiButton.Text = "Send to AI";
            sendAiButton.SetBounds(540, 370, 110, 24);
            sendAiButton.Click += (s, e) => StartChatRequest();

            runCmdButton = new Button();
            runCmdButton.Text = "Run Cmd";
            runCmdButton.SetBounds(660, 370, 110, 24);
            runCmdButton.Click += (s, e) => RunCommandFromInput();

            var endpointLabel = new Label();
            endpointLabel.Text = "Endpoint:";
            endpointLabel.SetBounds(10, 280, 70, 20);

            endpointBox = new TextBox();
            endpointBox.SetBounds(80, 278, 300, 22);

            var modelLabel = new Label();
            modelLabel.Text = "Model:";
            modelLabel.SetBounds(10, 310, 70, 20);

            modelBox = new TextBox();
            modelBox.SetBounds(80, 308, 300, 22);

            statusLabel = new Label();
            statusLabel.Text = "Ready";
            statusLabel.SetBounds(10, 370, 520, 20);

            Controls.Add(termBox);
            Controls.Add(chatBox);
            Controls.Add(inputBox);
            Controls.Add(sendAiButton);
            Controls.Add(runCmdButton);
            Controls.Add(endpointLabel);
            Controls.Add(endpointBox);
            Controls.Add(modelLabel);
            Controls.Add(modelBox);
            Controls.Add(statusLabel);

            LoadConfig();
            StartCmd();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            try { if (cmdIn != null) cmdIn.Close(); } catch { }
            try { if (cmdProc != null && !cmdProc.HasExited) cmdProc.Kill(); } catch { }
            try { if (cmdReaderThread != null && cmdReaderThread.IsAlive) cmdReaderThread.Join(200); } catch { }
        }

        private void LoadConfig()
        {
            var endpoint = new StringBuilder(256);
            var model = new StringBuilder(128);
            GetPrivateProfileString("ollama", "endpoint", DefaultEndpoint, endpoint, endpoint.Capacity, "config.ini");
            GetPrivateProfileString("ollama", "model", DefaultModel, model, model.Capacity, "config.ini");
            endpointBox.Text = endpoint.ToString();
            modelBox.Text = model.ToString();
        }

        private void SaveConfig()
        {
            WritePrivateProfileString("ollama", "endpoint", endpointBox.Text, "config.ini");
            WritePrivateProfileString("ollama", "model", modelBox.Text, "config.ini");
        }

        private void StartCmd()
        {
            try
            {
                var psi = new ProcessStartInfo("cmd.exe");
                psi.UseShellExecute = false;
                psi.RedirectStandardInput = true;
                psi.RedirectStandardOutput = true;
                psi.RedirectStandardError = true;
                psi.CreateNoWindow = true;
                cmdProc = Process.Start(psi);
                cmdIn = cmdProc.StandardInput;
                cmdReaderThread = new Thread(ReadCmdOutput);
                cmdReaderThread.IsBackground = true;
                cmdReaderThread.Start();
            }
            catch
            {
                AppendLine(termBox, "(failed to spawn cmd.exe)");
            }
        }

        private void ReadCmdOutput()
        {
            try
            {
                var reader = cmdProc.StandardOutput;
                var errReader = cmdProc.StandardError;
                char[] buf = new char[512];
                while (!cmdProc.HasExited)
                {
                    if (reader.Peek() > -1)
                    {
                        int read = reader.Read(buf, 0, buf.Length);
                        if (read > 0) AppendTextThreadSafe(termBox, new string(buf, 0, read));
                    }
                    if (errReader.Peek() > -1)
                    {
                        int read = errReader.Read(buf, 0, buf.Length);
                        if (read > 0) AppendTextThreadSafe(termBox, new string(buf, 0, read));
                    }
                    Thread.Sleep(10);
                }
            }
            catch { }
        }

        private void RunCommandFromInput()
        {
            var cmd = inputBox.Text.Trim();
            if (cmd.Length == 0 || cmdIn == null) return;
            cmdIn.WriteLine(cmd);
            cmdIn.Flush();
            AppendLine(chatBox, "Task:");
            AppendLine(chatBox, cmd);
            AppendLine(chatBox, "");
            inputBox.Text = "";
        }

        private void StartChatRequest()
        {
            var prompt = inputBox.Text.Trim();
            if (prompt.Length == 0) return;
            inputBox.Text = "";
            AppendLine(chatBox, "User:");
            AppendLine(chatBox, prompt);
            AppendLine(chatBox, "");
            SaveConfig();

            var thread = new Thread(() => ChatRequest(prompt));
            thread.IsBackground = true;
            thread.Start();
        }

        private void ChatRequest(string prompt)
        {
            SetStatus("Calling Ollama...");
            string endpoint = endpointBox.Text.Trim();
            string model = modelBox.Text.Trim();
            string url = endpoint.TrimEnd('/') + "/api/generate";
            string body = "{\"model\":\"" + EscapeJson(model) + "\",\"prompt\":\"" + EscapeJson(prompt) + "\",\"stream\":false}";

            try
            {
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "POST";
                req.ContentType = "application/json";
                byte[] data = Encoding.UTF8.GetBytes(body);
                using (var reqStream = req.GetRequestStream())
                {
                    reqStream.Write(data, 0, data.Length);
                }
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var reader = new StreamReader(resp.GetResponseStream()))
                {
                    string json = reader.ReadToEnd();
                    string answer = ExtractResponse(json);
                    AppendLineThreadSafe(chatBox, "AI:");
                    AppendLineThreadSafe(chatBox, string.IsNullOrEmpty(answer) ? "(no response)" : answer);
                    AppendLineThreadSafe(chatBox, "");
                    SetStatus("Ready");
                }
            }
            catch
            {
                AppendLineThreadSafe(chatBox, "AI:");
                AppendLineThreadSafe(chatBox, "(failed to reach Ollama)");
                AppendLineThreadSafe(chatBox, "");
                SetStatus("Ollama error");
            }
        }

        private static string EscapeJson(string s)
        {
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        }

        private static string ExtractResponse(string json)
        {
            const string key = "\"response\":\"";
            int start = json.IndexOf(key, StringComparison.Ordinal);
            if (start < 0) return "";
            start += key.Length;
            int end = start;
            while (end < json.Length)
            {
                if (json[end] == '"' && json[end - 1] != '\\') break;
                end++;
            }
            if (end <= start) return "";
            string raw = json.Substring(start, end - start);
            return UnescapeJson(raw);
        }

        private static string UnescapeJson(string s)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '\\' && i + 1 < s.Length)
                {
                    char c = s[i + 1];
                    if (c == 'n') { sb.Append('\n'); i++; continue; }
                    if (c == 'r') { sb.Append('\r'); i++; continue; }
                    if (c == 't') { sb.Append('\t'); i++; continue; }
                    if (c == '\\' || c == '"') { sb.Append(c); i++; continue; }
                }
                sb.Append(s[i]);
            }
            return sb.ToString();
        }

        private void AppendTextThreadSafe(RichTextBox box, string text)
        {
            if (box.InvokeRequired)
            {
                box.Invoke(new Action<RichTextBox, string>(AppendTextThreadSafe), box, text);
                return;
            }
            box.AppendText(text);
        }

        private void AppendLineThreadSafe(RichTextBox box, string text)
        {
            AppendTextThreadSafe(box, text + Environment.NewLine);
        }

        private static void AppendLine(RichTextBox box, string text)
        {
            box.AppendText(text + Environment.NewLine);
        }

        private void SetStatus(string text)
        {
            if (statusLabel.InvokeRequired)
            {
                statusLabel.Invoke(new Action<string>(SetStatus), text);
                return;
            }
            statusLabel.Text = text;
        }
    }
}
