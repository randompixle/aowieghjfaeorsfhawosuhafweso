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
        private TextBox xpHostBox;
        private TextBox xpPortBox;
        private Label statusLabel;
        private Button sendAiButton;
        private Button runCmdButton;
        private Button sendXpCmdButton;
        private CheckBox sendAiToXpCheck;

        private Process cmdProc;
        private StreamWriter cmdIn;
        private System.Collections.Generic.List<string> chatHistory = new System.Collections.Generic.List<string>();
        private const int MaxHistoryLines = 12;

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

            sendXpCmdButton = new Button();
            sendXpCmdButton.Text = "Send XP Cmd";
            sendXpCmdButton.SetBounds(780, 370, 130, 24);
            sendXpCmdButton.Click += (s, e) => SendXpCmdFromInput();

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

            var xpHostLabel = new Label();
            xpHostLabel.Text = "XP Host:";
            xpHostLabel.SetBounds(10, 340, 70, 20);

            xpHostBox = new TextBox();
            xpHostBox.SetBounds(80, 338, 200, 22);

            var xpPortLabel = new Label();
            xpPortLabel.Text = "XP Port:";
            xpPortLabel.SetBounds(290, 340, 60, 20);

            xpPortBox = new TextBox();
            xpPortBox.SetBounds(350, 338, 80, 22);

            sendAiToXpCheck = new CheckBox();
            sendAiToXpCheck.Text = "Send AI actions to XP";
            sendAiToXpCheck.SetBounds(540, 400, 200, 20);

            statusLabel = new Label();
            statusLabel.Text = "Ready";
            statusLabel.SetBounds(10, 370, 520, 20);

            Controls.Add(termBox);
            Controls.Add(chatBox);
            Controls.Add(inputBox);
            Controls.Add(sendAiButton);
            Controls.Add(runCmdButton);
            Controls.Add(sendXpCmdButton);
            Controls.Add(endpointLabel);
            Controls.Add(endpointBox);
            Controls.Add(modelLabel);
            Controls.Add(modelBox);
            Controls.Add(xpHostLabel);
            Controls.Add(xpHostBox);
            Controls.Add(xpPortLabel);
            Controls.Add(xpPortBox);
            Controls.Add(sendAiToXpCheck);
            Controls.Add(statusLabel);

            LoadConfig();
            StartCmd();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            try { if (cmdIn != null) cmdIn.Close(); } catch { }
            try { if (cmdProc != null && !cmdProc.HasExited) cmdProc.Kill(); } catch { }
        }

        private void LoadConfig()
        {
            var endpoint = new StringBuilder(256);
            var model = new StringBuilder(128);
            GetPrivateProfileString("ollama", "endpoint", DefaultEndpoint, endpoint, endpoint.Capacity, "config.ini");
            GetPrivateProfileString("ollama", "model", DefaultModel, model, model.Capacity, "config.ini");
            endpointBox.Text = endpoint.ToString();
            modelBox.Text = model.ToString();

            var xpHost = new StringBuilder(128);
            var xpPort = new StringBuilder(16);
            GetPrivateProfileString("xp", "host", "192.168.29.129", xpHost, xpHost.Capacity, "config.ini");
            GetPrivateProfileString("xp", "port", "6001", xpPort, xpPort.Capacity, "config.ini");
            xpHostBox.Text = xpHost.ToString();
            xpPortBox.Text = xpPort.ToString();
        }

        private void SaveConfig()
        {
            WritePrivateProfileString("ollama", "endpoint", endpointBox.Text, "config.ini");
            WritePrivateProfileString("ollama", "model", modelBox.Text, "config.ini");
            WritePrivateProfileString("xp", "host", xpHostBox.Text, "config.ini");
            WritePrivateProfileString("xp", "port", xpPortBox.Text, "config.ini");
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
                cmdProc = new Process();
                cmdProc.StartInfo = psi;
                cmdProc.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        AppendLineThreadSafe(termBox, e.Data);
                };
                cmdProc.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        AppendLineThreadSafe(termBox, e.Data);
                };
                cmdProc.Start();
                cmdIn = cmdProc.StandardInput;
                cmdProc.BeginOutputReadLine();
                cmdProc.BeginErrorReadLine();
            }
            catch
            {
                AppendLine(termBox, "(failed to spawn cmd.exe)");
            }
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

        private void SendXpCmdFromInput()
        {
            var line = inputBox.Text.Trim();
            if (line.Length == 0) return;
            SaveConfig();
            try
            {
                SendLinesToXp(new string[] { "cmd " + line });
                AppendLine(chatBox, "XP Cmd:");
                AppendLine(chatBox, line);
                AppendLine(chatBox, "");
                inputBox.Text = "";
                SetStatus("XP command sent");
            }
            catch
            {
                AppendLine(chatBox, "XP Cmd:");
                AppendLine(chatBox, "(failed to send)");
                AppendLine(chatBox, "");
                SetStatus("XP send error");
            }
        }

        private void StartChatRequest()
        {
            var prompt = inputBox.Text.Trim();
            if (prompt.Length == 0) return;
            inputBox.Text = "";
            AppendLine(chatBox, "User:");
            AppendLine(chatBox, prompt);
            AppendLine(chatBox, "");
            AddHistory("User: " + prompt);
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
            string systemHint =
                "You are integrated into a Windows XP app that provides a local cmd.exe terminal. " +
                "You can run commands by writing \"!cmd <cmd>\" in your response. " +
                "Do not claim you lack access to the terminal. " +
                "Prefer a dedicated line like: !cmd echo hi. " +
                "Only output task-relevant text.";
            string history = BuildHistory();
            string fullPrompt = systemHint + "\n\nRecent chat:\n" + history + "\n\nUser:\n" + prompt;
            string body = "{\"model\":\"" + EscapeJson(model) + "\",\"prompt\":\"" + EscapeJson(fullPrompt) + "\",\"stream\":false}";

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
                    string visible = HandleCmdDirectives(answer);
                    AppendLineThreadSafe(chatBox, string.IsNullOrEmpty(visible) ? "(no response)" : visible);
                    AppendLineThreadSafe(chatBox, "");
                    if (!string.IsNullOrEmpty(visible))
                        AddHistory("AI: " + visible);
                    if (sendAiToXpCheck.Checked && !string.IsNullOrEmpty(answer))
                    {
                        TrySendActionsToXp(answer);
                    }
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

        private string HandleCmdDirectives(string answer)
        {
            if (string.IsNullOrEmpty(answer)) return "";
            var lines = answer.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            var visible = new StringBuilder();
            foreach (var raw in lines)
            {
                var line = raw.TrimEnd();
                bool hadCmd = false;
                int idx = 0;
                while ((idx = line.IndexOf("!cmd ", idx, StringComparison.Ordinal)) >= 0)
                {
                    var cmd = line.Substring(idx + 5).Trim();
                    if (cmd.Length > 0 && cmdIn != null)
                    {
                        cmdIn.WriteLine(cmd);
                        cmdIn.Flush();
                    }
                    hadCmd = true;
                    break;
                }
                if (hadCmd) continue;
                if (visible.Length > 0) visible.AppendLine();
                visible.Append(line);
            }
            return visible.ToString().Trim();
        }

        private void AddHistory(string line)
        {
            chatHistory.Add(line);
            while (chatHistory.Count > MaxHistoryLines)
                chatHistory.RemoveAt(0);
        }

        private string BuildHistory()
        {
            if (chatHistory.Count == 0) return "(none)";
            return string.Join("\n", chatHistory.ToArray());
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

        private void SendLinesToXp(string[] lines)
        {
            string host = xpHostBox.Text.Trim();
            int port = 6001;
            int.TryParse(xpPortBox.Text.Trim(), out port);
            if (port <= 0) port = 6001;

            using (var client = new System.Net.Sockets.TcpClient())
            {
                client.ReceiveTimeout = 5000;
                client.SendTimeout = 5000;
                client.Connect(host, port);
                using (var stream = client.GetStream())
                {
                    string payload = string.Join("\n", lines) + "\n";
                    byte[] data = Encoding.UTF8.GetBytes(payload);
                    stream.Write(data, 0, data.Length);
                    stream.Flush();
                }
            }
        }

        private void TrySendActionsToXp(string text)
        {
            var lines = new System.Collections.Generic.List<string>();
            var parts = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            foreach (var raw in parts)
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;
                lines.Add(line);
            }
            if (lines.Count == 0) return;
            try
            {
                SendLinesToXp(lines.ToArray());
                SetStatus("AI actions sent to XP");
            }
            catch
            {
                SetStatus("AI send to XP failed");
            }
        }
    }
}
