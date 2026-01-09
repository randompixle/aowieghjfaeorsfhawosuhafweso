using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
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
        private RichTextBox actionLogBox;
        private RichTextBox scriptBox;
        private TextBox inputBox;
        private TextBox endpointBox;
        private TextBox modelBox;
        private TextBox xpHostBox;
        private TextBox xpPortBox;
        private Label statusLabel;
        private Label ollamaStatusLabel;
        private Label xpStatusLabel;
        private Button sendAiButton;
        private Button runCmdButton;
        private Button sendXpCmdButton;
        private Button runScriptButton;
        private Button checkConnButton;
        private Button clearMemoryButton;
        private Button runPresetButton;
        private CheckBox sendAiToXpCheck;
        private CheckBox dryRunCheck;
        private ComboBox presetBox;
        private NumericUpDown historySizeUpDown;
        private NumericUpDown retryCountUpDown;
        private TextBox allowListBox;
        private TextBox denyListBox;

        private Process cmdProc;
        private StreamWriter cmdIn;
        private System.Collections.Generic.List<string> chatHistory = new System.Collections.Generic.List<string>();
        private int maxHistoryLines = 12;

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
            Width = 980;
            Height = 650;
            StartPosition = FormStartPosition.CenterScreen;

            termBox = new RichTextBox();
            termBox.ReadOnly = true;
            termBox.Multiline = true;
            termBox.ScrollBars = RichTextBoxScrollBars.Vertical;
            termBox.SetBounds(10, 10, 520, 300);

            chatBox = new RichTextBox();
            chatBox.ReadOnly = true;
            chatBox.Multiline = true;
            chatBox.ScrollBars = RichTextBoxScrollBars.Vertical;
            chatBox.SetBounds(540, 10, 400, 260);

            inputBox = new TextBox();
            inputBox.Multiline = true;
            inputBox.ScrollBars = ScrollBars.Vertical;
            inputBox.SetBounds(540, 280, 400, 60);
            inputBox.KeyDown += InputBoxKeyDown;

            actionLogBox = new RichTextBox();
            actionLogBox.ReadOnly = true;
            actionLogBox.Multiline = true;
            actionLogBox.ScrollBars = RichTextBoxScrollBars.Vertical;
            actionLogBox.SetBounds(10, 320, 520, 120);

            scriptBox = new RichTextBox();
            scriptBox.Multiline = true;
            scriptBox.ScrollBars = RichTextBoxScrollBars.Vertical;
            scriptBox.SetBounds(540, 350, 400, 80);

            sendAiButton = new Button();
            sendAiButton.Text = "Send to AI";
            sendAiButton.SetBounds(540, 440, 110, 24);
            sendAiButton.Click += (s, e) => StartChatRequest();

            runCmdButton = new Button();
            runCmdButton.Text = "Run Cmd";
            runCmdButton.SetBounds(660, 440, 110, 24);
            runCmdButton.Click += (s, e) => RunCommandFromInput();

            sendXpCmdButton = new Button();
            sendXpCmdButton.Text = "Send XP Cmd";
            sendXpCmdButton.SetBounds(780, 440, 130, 24);
            sendXpCmdButton.Click += (s, e) => SendXpCmdFromInput();

            runScriptButton = new Button();
            runScriptButton.Text = "Run Script";
            runScriptButton.SetBounds(540, 470, 110, 24);
            runScriptButton.Click += (s, e) => RunScript();

            runPresetButton = new Button();
            runPresetButton.Text = "Run Preset";
            runPresetButton.SetBounds(660, 470, 110, 24);
            runPresetButton.Click += (s, e) => RunPreset();

            checkConnButton = new Button();
            checkConnButton.Text = "Check Conn";
            checkConnButton.SetBounds(780, 470, 130, 24);
            checkConnButton.Click += (s, e) => CheckConnections();

            var endpointLabel = new Label();
            endpointLabel.Text = "Endpoint:";
            endpointLabel.SetBounds(10, 450, 70, 20);

            endpointBox = new TextBox();
            endpointBox.SetBounds(80, 448, 300, 22);

            var modelLabel = new Label();
            modelLabel.Text = "Model:";
            modelLabel.SetBounds(10, 480, 70, 20);

            modelBox = new TextBox();
            modelBox.SetBounds(80, 478, 300, 22);

            var xpHostLabel = new Label();
            xpHostLabel.Text = "XP Host:";
            xpHostLabel.SetBounds(10, 510, 70, 20);

            xpHostBox = new TextBox();
            xpHostBox.SetBounds(80, 508, 200, 22);

            var xpPortLabel = new Label();
            xpPortLabel.Text = "XP Port:";
            xpPortLabel.SetBounds(290, 510, 60, 20);

            xpPortBox = new TextBox();
            xpPortBox.SetBounds(350, 508, 80, 22);

            sendAiToXpCheck = new CheckBox();
            sendAiToXpCheck.Text = "Send AI actions to XP";
            sendAiToXpCheck.SetBounds(540, 500, 200, 20);

            dryRunCheck = new CheckBox();
            dryRunCheck.Text = "Dry run";
            dryRunCheck.SetBounds(750, 500, 80, 20);

            var historyLabel = new Label();
            historyLabel.Text = "History:";
            historyLabel.SetBounds(10, 540, 60, 20);

            historySizeUpDown = new NumericUpDown();
            historySizeUpDown.Minimum = 2;
            historySizeUpDown.Maximum = 100;
            historySizeUpDown.Value = maxHistoryLines;
            historySizeUpDown.SetBounds(70, 538, 60, 22);
            historySizeUpDown.ValueChanged += (s, e) => maxHistoryLines = (int)historySizeUpDown.Value;

            clearMemoryButton = new Button();
            clearMemoryButton.Text = "Clear Memory";
            clearMemoryButton.SetBounds(140, 538, 110, 24);
            clearMemoryButton.Click += (s, e) => { chatHistory.Clear(); LogAction("Memory cleared."); };

            var retryLabel = new Label();
            retryLabel.Text = "Retry:";
            retryLabel.SetBounds(260, 540, 50, 20);

            retryCountUpDown = new NumericUpDown();
            retryCountUpDown.Minimum = 0;
            retryCountUpDown.Maximum = 5;
            retryCountUpDown.Value = 1;
            retryCountUpDown.SetBounds(310, 538, 40, 22);

            var allowLabel = new Label();
            allowLabel.Text = "Allow:";
            allowLabel.SetBounds(360, 540, 50, 20);

            allowListBox = new TextBox();
            allowListBox.SetBounds(410, 538, 120, 22);

            var denyLabel = new Label();
            denyLabel.Text = "Deny:";
            denyLabel.SetBounds(540, 540, 40, 20);

            denyListBox = new TextBox();
            denyListBox.SetBounds(580, 538, 120, 22);

            var presetLabel = new Label();
            presetLabel.Text = "Preset:";
            presetLabel.SetBounds(540, 560, 50, 20);

            presetBox = new ComboBox();
            presetBox.SetBounds(590, 558, 200, 22);
            presetBox.DropDownStyle = ComboBoxStyle.DropDownList;
            presetBox.Items.AddRange(new object[] { "dir", "ipconfig", "systeminfo", "whoami" });
            if (presetBox.Items.Count > 0) presetBox.SelectedIndex = 0;

            ollamaStatusLabel = new Label();
            ollamaStatusLabel.Text = "Ollama: ?";
            ollamaStatusLabel.SetBounds(10, 570, 150, 20);

            xpStatusLabel = new Label();
            xpStatusLabel.Text = "XP: ?";
            xpStatusLabel.SetBounds(170, 570, 150, 20);

            statusLabel = new Label();
            statusLabel.Text = "Ready";
            statusLabel.SetBounds(10, 590, 520, 20);

            Controls.Add(termBox);
            Controls.Add(chatBox);
            Controls.Add(inputBox);
            Controls.Add(actionLogBox);
            Controls.Add(scriptBox);
            Controls.Add(sendAiButton);
            Controls.Add(runCmdButton);
            Controls.Add(sendXpCmdButton);
            Controls.Add(runScriptButton);
            Controls.Add(runPresetButton);
            Controls.Add(checkConnButton);
            Controls.Add(endpointLabel);
            Controls.Add(endpointBox);
            Controls.Add(modelLabel);
            Controls.Add(modelBox);
            Controls.Add(xpHostLabel);
            Controls.Add(xpHostBox);
            Controls.Add(xpPortLabel);
            Controls.Add(xpPortBox);
            Controls.Add(sendAiToXpCheck);
            Controls.Add(dryRunCheck);
            Controls.Add(historyLabel);
            Controls.Add(historySizeUpDown);
            Controls.Add(clearMemoryButton);
            Controls.Add(retryLabel);
            Controls.Add(retryCountUpDown);
            Controls.Add(allowLabel);
            Controls.Add(allowListBox);
            Controls.Add(denyLabel);
            Controls.Add(denyListBox);
            Controls.Add(presetLabel);
            Controls.Add(presetBox);
            Controls.Add(ollamaStatusLabel);
            Controls.Add(xpStatusLabel);
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

            var history = new StringBuilder(16);
            GetPrivateProfileString("ui", "history_lines", "12", history, history.Capacity, "config.ini");
            int h;
            if (int.TryParse(history.ToString(), out h) && h >= 2 && h <= 100)
            {
                maxHistoryLines = h;
                historySizeUpDown.Value = h;
            }
            var retry = new StringBuilder(8);
            GetPrivateProfileString("ui", "retry_count", "1", retry, retry.Capacity, "config.ini");
            int r;
            if (int.TryParse(retry.ToString(), out r) && r >= 0 && r <= 5)
                retryCountUpDown.Value = r;

            var allow = new StringBuilder(128);
            var deny = new StringBuilder(128);
            GetPrivateProfileString("ui", "allow_cmds", "", allow, allow.Capacity, "config.ini");
            GetPrivateProfileString("ui", "deny_cmds", "", deny, deny.Capacity, "config.ini");
            allowListBox.Text = allow.ToString();
            denyListBox.Text = deny.ToString();
        }

        private void SaveConfig()
        {
            WritePrivateProfileString("ollama", "endpoint", endpointBox.Text, "config.ini");
            WritePrivateProfileString("ollama", "model", modelBox.Text, "config.ini");
            WritePrivateProfileString("xp", "host", xpHostBox.Text, "config.ini");
            WritePrivateProfileString("xp", "port", xpPortBox.Text, "config.ini");
            WritePrivateProfileString("ui", "history_lines", maxHistoryLines.ToString(), "config.ini");
            WritePrivateProfileString("ui", "retry_count", retryCountUpDown.Value.ToString(), "config.ini");
            WritePrivateProfileString("ui", "allow_cmds", allowListBox.Text, "config.ini");
            WritePrivateProfileString("ui", "deny_cmds", denyListBox.Text, "config.ini");
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
            if (!dryRunCheck.Checked)
            {
                cmdIn.WriteLine(cmd);
                cmdIn.Flush();
            }
            LogAction("Local cmd: " + cmd + (dryRunCheck.Checked ? " (dry run)" : ""));
            AppendLine(chatBox, "Task:");
            AppendLine(chatBox, cmd);
            AppendLine(chatBox, "");
            inputBox.Text = "";
        }

        private void InputBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                StartChatRequest();
                e.SuppressKeyPress = true;
                return;
            }
            if (e.Control && e.Shift && e.KeyCode == Keys.Enter)
            {
                RunCommandFromInput();
                e.SuppressKeyPress = true;
                return;
            }
            if (e.Control && e.Alt && e.KeyCode == Keys.Enter)
            {
                SendXpCmdFromInput();
                e.SuppressKeyPress = true;
            }
        }

        private void SendXpCmdFromInput()
        {
            var line = inputBox.Text.Trim();
            if (line.Length == 0) return;
            SaveConfig();
            try
            {
                if (!dryRunCheck.Checked)
                    SendLinesToXp(new string[] { "cmd " + line });
                LogAction("XP cmd: " + line + (dryRunCheck.Checked ? " (dry run)" : ""));
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

        private void RunScript()
        {
            var text = scriptBox.Text.Replace("\r\n", "\n").Replace("\r", "\n");
            var lines = text.Split('\n');
            foreach (var raw in lines)
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;
                if (line.StartsWith("!cmd "))
                {
                    HandleCmdLine(line);
                }
                else
                {
                    if (!dryRunCheck.Checked)
                        SendLinesToXp(new string[] { line });
                    LogAction("XP script: " + line + (dryRunCheck.Checked ? " (dry run)" : ""));
                }
            }
        }

        private void RunPreset()
        {
            if (presetBox.SelectedItem == null) return;
            inputBox.Text = presetBox.SelectedItem.ToString();
            RunCommandFromInput();
        }

        private void CheckConnections()
        {
            ollamaStatusLabel.Text = "Ollama: ...";
            xpStatusLabel.Text = "XP: ...";
            ThreadPool.QueueUserWorkItem(_ =>
            {
                bool okOllama = CheckOllama();
                bool okXp = CheckXp();
                if (ollamaStatusLabel.InvokeRequired)
                {
                    ollamaStatusLabel.Invoke(new Action(() => ollamaStatusLabel.Text = "Ollama: " + (okOllama ? "OK" : "FAIL")));
                    xpStatusLabel.Invoke(new Action(() => xpStatusLabel.Text = "XP: " + (okXp ? "OK" : "FAIL")));
                }
                else
                {
                    ollamaStatusLabel.Text = "Ollama: " + (okOllama ? "OK" : "FAIL");
                    xpStatusLabel.Text = "XP: " + (okXp ? "OK" : "FAIL");
                }
            });
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
                "Avoid refusal language; take action via !cmd when asked. " +
                "Only output task-relevant text.";
            string history = BuildHistory();
            string fullPrompt = systemHint + "\n\nRecent chat:\n" + history + "\n\nUser:\n" + prompt;
            string body = "{\"model\":\"" + EscapeJson(model) + "\",\"prompt\":\"" + EscapeJson(fullPrompt) + "\",\"stream\":false}";

            try
            {
                int retries = (int)retryCountUpDown.Value;
                for (int attempt = 0; attempt <= retries; attempt++)
                {
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
                            return;
                        }
                    }
                    catch
                    {
                        if (attempt >= retries) throw;
                        Thread.Sleep(400);
                    }
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
                if (HandleCmdLine(line)) continue;
                if (visible.Length > 0) visible.AppendLine();
                visible.Append(line);
            }
            return visible.ToString().Trim();
        }

        private bool HandleCmdLine(string line)
        {
            int idx = 0;
            bool handled = false;
            while ((idx = line.IndexOf("!cmd ", idx, StringComparison.Ordinal)) >= 0)
            {
                var cmd = line.Substring(idx + 5).Trim();
                if (cmd.Length > 0)
                {
                    if (IsCmdAllowed(cmd))
                    {
                        if (!dryRunCheck.Checked && cmdIn != null)
                        {
                            cmdIn.WriteLine(cmd);
                            cmdIn.Flush();
                        }
                        LogAction("AI !cmd: " + cmd + (dryRunCheck.Checked ? " (dry run)" : ""));
                    }
                    else
                    {
                        LogAction("Blocked !cmd: " + cmd);
                    }
                }
                handled = true;
                idx += 5;
            }
            return handled;
        }

        private bool IsCmdAllowed(string cmd)
        {
            string name = cmd.Split(' ')[0].Trim().ToLowerInvariant();
            var deny = SplitList(denyListBox.Text);
            foreach (var d in deny)
            {
                if (name == d) return false;
            }
            var allow = SplitList(allowListBox.Text);
            if (allow.Length == 0) return true;
            foreach (var a in allow)
            {
                if (name == a) return true;
            }
            return false;
        }

        private string[] SplitList(string text)
        {
            if (string.IsNullOrEmpty(text)) return new string[0];
            var parts = text.Split(new char[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
                parts[i] = parts[i].Trim().ToLowerInvariant();
            return parts;
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

        private void LogAction(string text)
        {
            string stamp = DateTime.Now.ToString("HH:mm:ss");
            AppendLine(actionLogBox, "[" + stamp + "] " + text);
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
                if (!dryRunCheck.Checked)
                    SendLinesToXp(lines.ToArray());
                LogAction("AI actions -> XP (" + lines.Count + ")" + (dryRunCheck.Checked ? " (dry run)" : ""));
                SetStatus("AI actions sent to XP");
            }
            catch
            {
                SetStatus("AI send to XP failed");
            }
        }

        private bool CheckOllama()
        {
            try
            {
                string endpoint = endpointBox.Text.Trim();
                string url = endpoint.TrimEnd('/') + "/api/tags";
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "GET";
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    return resp.StatusCode == HttpStatusCode.OK;
                }
            }
            catch
            {
                return false;
            }
        }

        private bool CheckXp()
        {
            try
            {
                string host = xpHostBox.Text.Trim();
                int port = 6001;
                int.TryParse(xpPortBox.Text.Trim(), out port);
                if (port <= 0) port = 6001;
                using (var client = new TcpClient())
                {
                    client.ReceiveTimeout = 2000;
                    client.SendTimeout = 2000;
                    client.Connect(host, port);
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
