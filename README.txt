XP Ollama Terminal (.NET 4.0, XP 32-bit)

What it does
- Embeds a terminal view by running cmd.exe and streaming its output.
- Provides a chat pane to call Ollama (/api/generate) and show responses.
- Lets you run commands from the input box or send prompts to the model.

Build (on XP or a Windows machine with .NET 4.0 tools)
- Create a C# WinForms project (.NET Framework 4.0).
- Add Program.cs.
- Build for x86.
Or compile directly with csc:
C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /target:winexe /platform:x86 /out:XPAgent.exe ^
 /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll ^
 Program.cs

Runtime
- Edit config.ini to point to your Ollama host IP and model.
- If the XP VM cannot reach 127.0.0.1, use the host IP (e.g. http://192.168.56.1:11434).

Notes
- The JSON parsing is minimal and expects a non-streaming Ollama response.
- The input box is shared for "Run Cmd" and "Send to AI".
- XP command sending uses the guest agent command socket (default port 6001).
- The AI can run local cmd.exe commands by outputting "!cmd <command>".
- Use "Run Script" to send multiple action lines to the XP guest (prefix local commands with !cmd).
