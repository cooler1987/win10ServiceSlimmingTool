using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Security.Principal;
using System.ServiceProcess;
using System.Windows.Forms;
using Microsoft.Win32;

namespace WinServiceOptimizer
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            if (!IsAdministrator())
            {
                MessageBox.Show("请以【管理员身份】运行此程序，否则无法修改系统服务状态！", "权限不足", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Application.Run(new MainForm());
        }

        static bool IsAdministrator()
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }

    public class ServiceMeta
    {
        public string Category { get; set; } = "";
        public string Description { get; set; } = "";
        public Color BadgeColor { get; set; } = Color.Gray;
    }

    public class MainForm : Form
    {
        private ListView listView;
        private Button btnToggle;
        private Button btnRefresh;
        private CheckBox chkOnlyRunning;
        private Label lblStatus;

        // 预设服务分类知识库（涵盖常见系统服务及第三方服务）
        private static readonly Dictionary<string, ServiceMeta> KnowledgeBase = new(StringComparer.OrdinalIgnoreCase)
        {
            // 1. 推荐安全禁用
            ["DPS"] = new() { Category = "1. 推荐禁用", Description = "诊断策略服务：占用内存与磁盘I/O，禁用仅影响系统自带疑难解答向导。", BadgeColor = Color.SeaGreen },
            ["TrkWks"] = new() { Category = "1. 推荐禁用", Description = "分布式链接跟踪：局域网移动文件后修复快捷方式，单机完全无用。", BadgeColor = Color.SeaGreen },
            ["QWAVE"] = new() { Category = "1. 推荐禁用", Description = "音视频流媒体QoS调度：局域网流媒体优先级优化，日常无需开启。", BadgeColor = Color.SeaGreen },
            ["CDPSvc"] = new() { Category = "1. 推荐禁用", Description = "连接设备平台服务：用于跨设备同步、手机连接，不用多设备协同可禁用。", BadgeColor = Color.SeaGreen },
            ["CDPUserSvc"] = new() { Category = "1. 推荐禁用", Description = "连接设备平台用户服务：跨设备体验组件（已适配注册表彻底禁用）。", BadgeColor = Color.SeaGreen },
            ["WpnUserService"] = new() { Category = "1. 推荐禁用", Description = "系统推送通知用户服务：禁用后不再接收右下角UWP应用及系统弹窗推送。", BadgeColor = Color.SeaGreen },
            ["TabletInputService"] = new() { Category = "1. 推荐禁用", Description = "触摸键盘和手写面板：台式机或非触屏笔记本可安全禁用。", BadgeColor = Color.SeaGreen },
            ["Browser"] = new() { Category = "1. 推荐禁用", Description = "Computer Browser：维护局域网老旧SMBv1计算机列表，现代网络已淘汰。", BadgeColor = Color.SeaGreen },

            // 2. 按需禁用
            ["Spooler"] = new() { Category = "2. 按需禁用", Description = "打印后台处理程序：不使用实体打印机及虚拟PDF打印功能时可禁用。", BadgeColor = Color.DarkOrange },
            ["BthAvctpSvc"] = new() { Category = "2. 按需禁用", Description = "AVCTP蓝牙音视频控制：不使用蓝牙耳机或蓝牙音箱时可禁用。", BadgeColor = Color.DarkOrange },
            ["WbioSrvc"] = new() { Category = "2. 按需禁用", Description = "Windows生物识别服务：无指纹识别、无Windows Hello人脸解锁可禁用。", BadgeColor = Color.DarkOrange },
            ["WlanSvc"] = new() { Category = "2. 按需禁用", Description = "WLAN无线配置服务：纯插网线的台式机且永远不用Wi-Fi时可禁用。", BadgeColor = Color.DarkOrange },
            ["LanmanServer"] = new() { Category = "2. 按需禁用", Description = "Server文件共享服务：不需要让局域网其他电脑访问本机共享文件夹时可禁用。", BadgeColor = Color.DarkOrange },
            ["lmhosts"] = new() { Category = "2. 按需禁用", Description = "TCP/IP NetBIOS Helper：辅助局域网共享与名称解析，不用局域网共享可禁用。", BadgeColor = Color.DarkOrange },
            ["RasMan"] = new() { Category = "2. 按需禁用", Description = "远程访问连接管理器：不用系统自带VPN客户端或宽带拨号时可禁用。", BadgeColor = Color.DarkOrange },
            ["SstpSvc"] = new() { Category = "2. 按需禁用", Description = "SSTP VPN协议服务：不使用系统自带SSTP VPN时可禁用（不影响第三方代理）。", BadgeColor = Color.DarkOrange },
            ["camsvc"] = new() { Category = "2. 按需禁用", Description = "功能访问管理器：控制UWP应用对相机/麦克风的权限，不用UWP应用可考虑禁用。", BadgeColor = Color.DarkOrange },
            ["TokenBroker"] = new() { Category = "2. 按需禁用", Description = "Web帐户管理器：用于微软商店与UWP应用账号登录，纯本地账号不使用商店可禁用。", BadgeColor = Color.DarkOrange },

            // 3. 建议保留/核心系统
            ["AudioEndpointBuilder"] = new() { Category = "3. 核心勿动", Description = "音频端点生成器：禁用会导致电脑完全无声音。", BadgeColor = Color.Firebrick },
            ["Audiosrv"] = new() { Category = "3. 核心勿动", Description = "Windows音频服务：声卡核心服务，禁用会导致无声音。", BadgeColor = Color.Firebrick },
            ["BFE"] = new() { Category = "3. 核心勿动", Description = "基本筛选引擎：防火墙及网络过滤基石，禁用导致网络异常。", BadgeColor = Color.Firebrick },
            ["mpssvc"] = new() { Category = "3. 核心勿动", Description = "Windows Defender 防火墙：禁用会导致部分系统功能及商店异常。", BadgeColor = Color.Firebrick },
            ["Dhcp"] = new() { Category = "3. 核心勿动", Description = "DHCP客户端：自动获取IP地址，禁用会导致无法上网。", BadgeColor = Color.Firebrick },
            ["Dnscache"] = new() { Category = "3. 核心勿动", Description = "DNS客户端：域名解析缓存，禁用会导致网页无法打开。", BadgeColor = Color.Firebrick },
            ["nsi"] = new() { Category = "3. 核心勿动", Description = "网络存储接口服务：核心网络组件，禁用直接断网。", BadgeColor = Color.Firebrick },
            ["NlaSvc"] = new() { Category = "3. 核心勿动", Description = "网络位置感知：检测网络状态，禁用会导致任务栏网络图标异常。", BadgeColor = Color.Firebrick },
            ["netprofm"] = new() { Category = "3. 核心勿动", Description = "网络列表服务：识别已连接的网络，禁用会导致网络配置异常。", BadgeColor = Color.Firebrick },
            ["Wcmsvc"] = new() { Category = "3. 核心勿动", Description = "Windows连接管理器：管理网络切换与连接状态。", BadgeColor = Color.Firebrick },
            ["LanmanWorkstation"] = new() { Category = "3. 核心勿动", Description = "Workstation客户端：用于访问局域网共享及部分网络映射，建议保留。", BadgeColor = Color.Firebrick },
            ["hidserv"] = new() { Category = "3. 核心勿动", Description = "人体学接口设备服务：支持键盘多媒体快捷键、鼠标侧键等。", BadgeColor = Color.Firebrick },
            ["ShellHWDetection"] = new() { Category = "3. 核心勿动", Description = "外壳硬件检测：U盘、移动硬盘插入自动识别与通知。", BadgeColor = Color.Firebrick },
            ["Schedule"] = new() { Category = "3. 核心勿动", Description = "任务计划程序：系统维护与定时任务核心，禁用会导致多种故障。", BadgeColor = Color.Firebrick },
            ["Winmgmt"] = new() { Category = "3. 核心勿动", Description = "WMI系统管理规范：系统信息查询与驱动管理工具依赖。", BadgeColor = Color.Firebrick },
            ["EventLog"] = new() { Category = "3. 核心勿动", Description = "Windows事件日志：系统日志记录，许多服务依赖它启动。", BadgeColor = Color.Firebrick },
            ["CryptSvc"] = new() { Category = "3. 核心勿动", Description = "加密服务：验证驱动签名与系统证书，禁用会导致安装软件报错。", BadgeColor = Color.Firebrick },
            ["SepMasterService"] = new() { Category = "3. 核心勿动", Description = "Symantec杀毒主服务：受杀毒软件自我保护锁定，无法在此修改。", BadgeColor = Color.Firebrick },
            ["SepScanService"] = new() { Category = "3. 核心勿动", Description = "Symantec扫描服务：受杀毒软件自我保护锁定。", BadgeColor = Color.Firebrick },
            ["sepWscSvc"] = new() { Category = "3. 核心勿动", Description = "Symantec安全中心服务：受杀毒软件自我保护锁定。", BadgeColor = Color.Firebrick },
            ["RpcSs"] = new() { Category = "3. 核心勿动", Description = "远程过程调用(RPC)：Windows最底层通信核心，禁用即崩溃。", BadgeColor = Color.Firebrick },
            ["DcomLaunch"] = new() { Category = "3. 核心勿动", Description = "DCOM服务进程启动器：底层核心，禁用无法开机。", BadgeColor = Color.Firebrick },
            ["LSM"] = new() { Category = "3. 核心勿动", Description = "本地会话管理器：管理用户登录会话，禁用黑屏。", BadgeColor = Color.Firebrick },
            ["SamSs"] = new() { Category = "3. 核心勿动", Description = "安全账户管理器：管理本地账号密码，禁用无法登录。", BadgeColor = Color.Firebrick },
            ["ProfSvc"] = new() { Category = "3. 核心勿动", Description = "用户配置文件服务：加载桌面与个人设置，禁用无法进入桌面。", BadgeColor = Color.Firebrick },
            ["UserManager"] = new() { Category = "3. 核心勿动", Description = "用户管理器：多用户与会话管理核心。", BadgeColor = Color.Firebrick },
            ["Power"] = new() { Category = "3. 核心勿动", Description = "电源管理服务：控制睡眠、关机与CPU电源策略。", BadgeColor = Color.Firebrick },
            ["PlugPlay"] = new() { Category = "3. 核心勿动", Description = "即插即用服务：硬件设备识别核心。", BadgeColor = Color.Firebrick },
            ["StateRepository"] = new() { Category = "3. 核心勿动", Description = "状态存储库服务：开始菜单与现代应用模型依赖，禁用导致开始菜单卡死。", BadgeColor = Color.Firebrick }
        };

        public MainForm()
        {
            Text = "Windows 10 服务精简开关工具 (管理员模式)";
            Size = new Size(1080, 680);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9.5f);

            var topPanel = new Panel { Dock = DockStyle.Top, Height = 52, Padding = new Padding(10) };

            btnToggle = new Button
            {
                Text = "切换选中服务状态 (禁用 / 恢复启用)",
                Width = 260,
                Height = 32,
                Location = new Point(10, 10),
                BackColor = Color.RoyalBlue,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnToggle.Click += BtnToggle_Click;

            btnRefresh = new Button
            {
                Text = "刷新列表",
                Width = 95,
                Height = 32,
                Location = new Point(280, 10)
            };
            btnRefresh.Click += (s, e) => LoadServices();

            chkOnlyRunning = new CheckBox
            {
                Text = "仅显示【正在运行】的服务",
                Checked = true,
                AutoSize = true,
                Location = new Point(395, 16)
            };
            chkOnlyRunning.CheckedChanged += (s, e) => LoadServices();

            topPanel.Controls.AddRange(new Control[] { btnToggle, btnRefresh, chkOnlyRunning });

            var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 30, Padding = new Padding(10, 5, 10, 5) };
            lblStatus = new Label
            {
                Text = "提示：绿色为推荐安全禁用，橙色为按需禁用，红色为系统核心勿动（双击列表项也可切换）。",
                Dock = DockStyle.Fill,
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };
            bottomPanel.Controls.Add(lblStatus);

            listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = false
            };

            listView.Columns.Add("优化分类", 115);
            listView.Columns.Add("服务名称", 170);
            listView.Columns.Add("当前状态", 85);
            listView.Columns.Add("启动类型", 85);
            listView.Columns.Add("显示名称", 210);
            listView.Columns.Add("功能与禁用影响说明", 380);

            listView.DoubleClick += BtnToggle_Click;

            Controls.Add(listView);
            Controls.Add(bottomPanel);
            Controls.Add(topPanel);

            LoadServices();
        }

        private ServiceMeta GetMeta(string serviceName)
        {
            if (KnowledgeBase.TryGetValue(serviceName, out var meta))
                return meta;

            // 匹配带随机后缀的用户服务，如 CDPUserSvc_33c27
            int idx = serviceName.LastIndexOf('_');
            if (idx > 0)
            {
                string prefix = serviceName[..idx];
                if (KnowledgeBase.TryGetValue(prefix, out var prefixMeta))
                    return prefixMeta;
            }

            return new ServiceMeta
            {
                Category = "4. 其他系统服务",
                Description = "系统常规或第三方组件服务，若不确定用途建议保持默认。",
                BadgeColor = Color.DimGray
            };
        }

        private void LoadServices()
        {
            listView.BeginUpdate();
            listView.Items.Clear();

            var services = ServiceController.GetServices();
            var rows = new List<(ServiceMeta meta, string serviceName, string displayName, bool isRunning, string startMode)>();

            foreach (var sc in services)
            {
                using (sc)
                {
                    bool isRunning = sc.Status == ServiceControllerStatus.Running;
                    if (chkOnlyRunning.Checked && !isRunning)
                        continue;

                    var meta = GetMeta(sc.ServiceName);
                    string startMode = GetStartModeText(sc.ServiceName);
                    rows.Add((meta, sc.ServiceName, sc.DisplayName, isRunning, startMode));
                }
            }

            // 按分类优先级排序：1.推荐禁用 -> 2.按需禁用 -> 4.其他 -> 3.核心勿动
            var sorted = rows.OrderBy(r => r.meta.Category.StartsWith("3.") ? "9." : r.meta.Category)
                             .ThenBy(r => r.serviceName);

            foreach (var item in sorted)
            {
                var lvi = new ListViewItem(item.meta.Category);
                lvi.ForeColor = item.meta.BadgeColor;
                lvi.SubItems.Add(item.serviceName);
                lvi.SubItems.Add(item.isRunning ? "正在运行" : "已停止");
                lvi.SubItems.Add(item.startMode);
                lvi.SubItems.Add(item.displayName);
                lvi.SubItems.Add(item.meta.Description);
                lvi.Tag = item.serviceName;
                listView.Items.Add(lvi);
            }

            listView.EndUpdate();
        }

        private string GetStartModeText(string serviceName)
        {
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{serviceName}");
                if (key != null && key.GetValue("Start") is int startVal)
                {
                    return startVal switch
                    {
                        2 => "自动",
                        3 => "手动",
                        4 => "已禁用",
                        _ => "其他"
                    };
                }
            }
            catch { }
            return "未知";
        }

        private void BtnToggle_Click(object? sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
            {
                MessageBox.Show("请先在列表中点击选择一个服务！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selected = listView.SelectedItems[0];
            string category = selected.Text;
            string serviceName = selected.Tag?.ToString() ?? "";
            string currentStart = selected.SubItems[3].Text;

            if (category.StartsWith("3."))
            {
                var warn = MessageBox.Show(
                    $"【{serviceName}】属于系统核心或杀毒软件保护服务，强行禁用可能导致无法开机、断网或系统崩溃！\n\n确定仍要继续操作吗？",
                    "高危操作警告", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
                if (warn != DialogResult.Yes) return;
            }

            bool toDisable = currentStart != "已禁用";
            try
            {
                SetServiceState(serviceName, toDisable);
                lblStatus.Text = $"已成功将 [{serviceName}] 设为：{(toDisable ? "已禁用并停止" : "手动启用")}";
                LoadServices();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"修改服务 [{serviceName}] 失败：\n{ex.Message}\n\n（注：部分受系统或杀毒软件保护的服务无法直接更改）",
                    "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetServiceState(string serviceName, bool disable)
        {
            int targetStart = disable ? 4 : 3; // 4 = Disabled, 3 = Manual

            // 1. 写入注册表（同时兼容常规服务和带后缀的 Per-User 服务）
            using (var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{serviceName}", true))
            {
                key?.SetValue("Start", targetStart, RegistryValueKind.DWord);
            }

            // 如果是带后缀的用户服务，同步修改其模板项，防止重启后再生
            int idx = serviceName.LastIndexOf('_');
            if (idx > 0)
            {
                string baseName = serviceName[..idx];
                using var baseKey = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{baseName}", true);
                baseKey?.SetValue("Start", targetStart, RegistryValueKind.DWord);
            }

            // 2. 通过 ServiceController 停止或启动服务
            using var sc = new ServiceController(serviceName);
            if (disable && sc.Status != ServiceControllerStatus.Stopped && sc.CanStop)
            {
                sc.Stop();
                sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(5));
            }
            else if (!disable && sc.Status == ServiceControllerStatus.Stopped)
            {
                sc.Start();
                sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(5));
            }
        }
    }
}
