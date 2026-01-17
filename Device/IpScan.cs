using PoEWizard.Comm;
using PoEWizard.Data;
using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace PoEWizard.Device
{
    public static class IpScan
    {
        private const string PING_SCRIPT = "PoEWizard.Resources." + Constants.HELPER_SCRIPT_FILE_NAME;
        private const string REM_PATH = Constants.PYTHON_DIR + Constants.HELPER_SCRIPT_FILE_NAME;
        private const int SCAN_TIMEOUT = 290 * 1000;
        private const int PORT_TIMEOUT = 2500;
        private static SwitchModel model;
        private static SftpService sftpService;
        private static AosSshService sshService;

        public static void Init(SwitchModel model)
        {
            IpScan.model = model;
            sshService = new AosSshService(model);
            sftpService = new SftpService(model.IpAddress, model.Login, model.Password);
            sshService.ConnectSshClient();
            sftpService.Connect();
        }

        public static void Disconnect()
        {
            sshService.DisconnectSshClient();
            sftpService.Disconnect();
        }

        public async static Task LaunchScan()
        {
            try
            {
                using (var resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(PING_SCRIPT))
                {
                    if (sftpService.IsConnected)
                    {
                        long filesize = 0L;
                        int count = 0;
                        Logger.Activity($"Uploading python script to {REM_PATH} on {model.Name}");
                        sftpService.UploadStream(resource, REM_PATH, true);
                        while (filesize < resource.Length && count < 3)
                        {
                            Thread.Sleep(2000);
                            filesize = sftpService.GetFileSize(REM_PATH);
                            count++;
                        }
                        if (filesize == resource.Length)
                            Logger.Activity($"Uploading complete, filesize: {filesize / 1024} KB");
                        else
                            Logger.Error("Failed to upload python script to switch: timeout");
                        await RunScript();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error running IP address scan", ex);
            }
        }

        //Modified by SRH to eliminate the need for su mode but still verify if the helper script is running or not.
        //Note this depends on the udpated script that creates and response on the IPC socket.
        public static async Task<bool> IsIpScanRunning()
        {
            return await Task.Run(() =>
            {
                string Trunc(string s, int max)
                {
                    if (s == null) return "";
                    // Make log-friendly and visible
                    s = s.Replace("\r", "\\r").Replace("\n", "\\n");
                    return (s.Length <= max) ? s : s.Substring(0, max) + "...";
                }

                // -  Exits 0/1 so humans can test from CLI if needed.
                const string sockPath = "/dev/shm/installers_toolkit_helper.sock";
                string cmd =
                    "python3 -c 'import socket,sys; s=socket.socket(socket.AF_UNIX); s.settimeout(.5); " +
                    "r=s.connect_ex(\"" + sockPath + "\"); r==0 and sys.stdout.write(\"RUNNING\\n\"); " +
                    "sys.exit(0 if r==0 else 1)'";

                // MaxWaitSec must cover the 0.5s socket timeout plus SSH jitter + parsing
                // You were seeing ~2-4 sec failures; give it breathing room.
                var probeCmd = new LinuxCommand(cmd, null, 8);

                try
                {
                    Logger.Debug($"IsIpScanRunning: IPC probe starting (cmdLen={cmd.Length}, maxWaitSec={probeCmd.MaxWaitSec})");

                    Dictionary<string, string> resp = sshService.SendLinuxCommand(probeCmd);

                    string output = "";
                    if (resp != null && resp.ContainsKey("output") && resp["output"] != null)
                        output = resp["output"];

                    bool running = output.IndexOf("RUNNING", StringComparison.OrdinalIgnoreCase) >= 0;

                    Logger.Debug(
                        "IsIpScanRunning: IPC probe completed " +
                        $"(running={running}, outLen={output.Length}, outHead=\"{Trunc(output, 300)}\")");

                    return running;
                }
                catch (Exception ex)
                {
                    // This is the critical troubleshooting signal:
                    // if it still times out, we know the command echo is still being altered.
                    Logger.Error($"IsIpScanRunning: IPC probe threw exception: {ex.Message}", ex);
                    return false;
                }
            });
        }

        public static int GetOpenPort(string host)
        {

            foreach (int port in Constants.portsToScan)
            {
                Logger.Trace($"Begin checking host {host} port {port}");

                if (IsPortOpen(host, port))
                {
                    Logger.Trace($"{host}:{port} is open");
                    return port;
                }
                else
                {
                    Logger.Trace($"{host}:{port} is not open");
                }
            }
            return 0;
        }

        public static bool IsPortOpen(string host, int port)
        {
            if (string.IsNullOrEmpty(host)) return false;
            try
            {
                using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp) { Blocking = true })
                {
                    var result = socket.BeginConnect(host, port, null, null);
                    bool success = result.AsyncWaitHandle.WaitOne(PORT_TIMEOUT, false);
                    if (success) socket.EndConnect(result);
                    socket.Close();
                    return success;
                }
            }
            catch (SocketException)
            {
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error("Error checking port open", ex);
                return false;
            }
        }

        private async static Task RunScript()
        {
            try
            {
                Logger.Activity($"Launching python script on {model.Name}");
                Task<List<Dictionary<string, string>>> task = Task<List<Dictionary<string, string>>>.Factory.StartNew(() =>
                {
                    Dictionary<string, string> resp = sshService.SendCommand(new RestUrlEntry(Command.RUN_PYTHON_SCRIPT, new string[] { REM_PATH }), SCAN_TIMEOUT);
                    if (resp["output"].Contains("Err")) Logger.Error($"Failed to run ip scan on switch {model.Name}: {resp["output"]}");
                    resp = sshService.SendCommand(new RestUrlEntry(Command.SHOW_ARP), null);
                    return CliParseUtils.ParseHTable(resp["output"]);
                });
                if (await Task.WhenAny(task, Task.Delay(SCAN_TIMEOUT)) == task)
                {
                    List<Dictionary<string, string>> arp = task.Result;
                    model.LoadIPAdrressFromList(arp);
                }
                else
                {
                    Logger.Error($"Timeout waiting for ip address scan on switch {model.Name}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Could not scan IP addresses on switch {model.Name}", ex);
            }
        }
    }
}
