using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace WinServiceOptimizer
{
    static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDpiAwarenessContext(IntPtr value);

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main()
        {
            EnableDpiAwareness();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (!IsAdministrator())
            {
                MessageBox.Show("请以【管理员身份】运行此程序，否则无法修改系统服务状态！", "权限不足", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Application.Run(new MainForm());
        }

        static void EnableDpiAwareness()
        {
            try
            {
                // -4 代表 PerMonitorV2（Win10 1703+ 支持），失败或老系统自动降级
                if (!SetProcessDpiAwarenessContext(new IntPtr(-4)))
                    SetProcessDPIAware();
            }
            catch
            {
                try { SetProcessDPIAware(); } catch { }
            }
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

    public class ServiceRowInfo
    {
        public string Category { get; set; } = "";
        public Color BadgeColor { get; set; } = Color.DimGray;
        public string ServiceName { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public bool IsRunning { get; set; }
        public int StartModeValue { get; set; } // 2=Auto, 3=Manual, 4=Disabled
        public bool DelayedAutoStart { get; set; }
        public string StartModeText { get; set; } = "未知";
        public string SourceLabel { get; set; } = "未知";
        public bool IsThirdParty { get; set; }
        public string ExePath { get; set; } = "";
        public string RawImagePath { get; set; } = "";
        public int ProcessId { get; set; }
        public double MemoryMb { get; set; }
        public bool IsSharedHost { get; set; }
        public string MemoryText { get; set; } = "-";
        public string OfficialDescription { get; set; } = "";
        public string AdviceDescription { get; set; } = "";
        public List<string> DependsOn { get; set; } = new List<string>();
        public List<string> DependedBy { get; set; } = new List<string>();
    }

    [DataContract]
    public class ServiceBackupEntry
    {
        [DataMember] public string ServiceName { get; set; } = "";
        [DataMember] public string DisplayName { get; set; } = "";
        [DataMember] public int StartMode { get; set; }
        [DataMember] public int DelayedAutostart { get; set; }
        [DataMember] public bool WasRunning { get; set; }
    }

    [DataContract]
    public class BackupSnapshot
    {
        [DataMember] public string CreatedAt { get; set; } = "";
        [DataMember] public string MachineName { get; set; } = "";
        [DataMember] public List<ServiceBackupEntry> Services { get; set; } = new List<ServiceBackupEntry>();
    }

    public class BufferedListView : ListView
    {
        private bool _suppressCheckOnDblClick;

        public BufferedListView()
        {
            DoubleBuffered = true;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && e.Clicks >= 2)
                _suppressCheckOnDblClick = true;
            base.OnMouseDown(e);
        }

        protected override void OnItemCheck(ItemCheckEventArgs ice)
        {
            if (_suppressCheckOnDblClick)
            {
                ice.NewValue = ice.CurrentValue;
                _suppressCheckOnDblClick = false;
                return;
            }
            base.OnItemCheck(ice);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            _suppressCheckOnDblClick = false;
            base.OnMouseUp(e);
        }
    }

    public class ListViewColumnSorter : IComparer
    {
        public int SortColumn { get; set; } = 0;
        public SortOrder Order { get; set; } = SortOrder.Ascending;

        public int Compare(object x, object y)
        {
            var itemA = (ListViewItem)x;
            var itemB = (ListViewItem)y;
            var infoA = itemA.Tag as ServiceRowInfo;
            var infoB = itemB.Tag as ServiceRowInfo;

            int result;
            if (SortColumn == 5 && infoA != null && infoB != null)
            {
                // 内存列按数值排序
                result = infoA.MemoryMb.CompareTo(infoB.MemoryMb);
            }
            else if (SortColumn == 0 && infoA != null && infoB != null)
            {
                // 分类列按 1 -> 2 -> 4 -> 3(核心) 优先级排序
                string keyA = infoA.Category.StartsWith("3.") ? "9." : infoA.Category;
                string keyB = infoB.Category.StartsWith("3.") ? "9." : infoB.Category;
                result = string.Compare(keyA, keyB, StringComparison.OrdinalIgnoreCase);
                if (result == 0)
                    result = string.Compare(infoA.ServiceName, infoB.ServiceName, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                string textA = itemA.SubItems.Count > SortColumn ? itemA.SubItems[SortColumn].Text : "";
                string textB = itemB.SubItems.Count > SortColumn ? itemB.SubItems[SortColumn].Text : "";
                result = string.Compare(textA, textB, StringComparison.OrdinalIgnoreCase);
            }

            return Order == SortOrder.Descending ? -result : result;
        }
    }

    public class MainForm : Form
    {
        #region Win32 API

        [DllImport("shlwapi.dll", BestFitMapping = false, CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = false)]
        private static extern int SHLoadIndirectString(string pszSource, StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr OpenSCManager(string machineName, string databaseName, uint dwAccess);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr OpenService(IntPtr hSCManager, string lpServiceName, uint dwDesiredAccess);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool CloseServiceHandle(IntPtr hSCObject);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool QueryServiceStatusEx(IntPtr hService, int infoLevel, IntPtr lpBuffer, int cbBufSize, out int pcbBytesNeeded);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool ChangeServiceConfig(
            IntPtr hService, uint dwServiceType, uint dwStartType, uint dwErrorControl,
            string lpBinaryPathName, string lpLoadOrderGroup, IntPtr lpdwTagId,
            string lpDependencies, string lpServiceStartName, string lpPassword, string lpDisplayName);

        [StructLayout(LayoutKind.Sequential)]
        private struct SERVICE_STATUS_PROCESS
        {
            public int dwServiceType;
            public int dwCurrentState;
            public int dwControlsAccepted;
            public int dwWin32ExitCode;
            public int dwServiceSpecificExitCode;
            public int dwCheckPoint;
            public int dwWaitHint;
            public int dwProcessId;
            public int dwServiceFlags;
        }

        private const uint SC_MANAGER_CONNECT = 0x0001;
        private const uint SERVICE_QUERY_STATUS = 0x0004;
        private const uint SERVICE_CHANGE_CONFIG = 0x0002;
        private const uint SERVICE_NO_CHANGE = 0xFFFFFFFF;
        private const int SC_STATUS_PROCESS_INFO = 0;

        #endregion

        private readonly string _baseDir = AppDomain.CurrentDomain.BaseDirectory;
        private readonly string _initialBackupPath;
        private readonly string _logFilePath;
        private readonly List<string> _sessionLogs = new List<string>();
        private Dictionary<string, ServiceBackupEntry> _initialSnapshotMap = new Dictionary<string, ServiceBackupEntry>(StringComparer.OrdinalIgnoreCase);

        private List<ServiceRowInfo> _allServices = new List<ServiceRowInfo>();
        private readonly ListViewColumnSorter _sorter = new ListViewColumnSorter();
        private bool _isBusy;

        private BufferedListView listView;
        private TextBox txtSearch;
        private ComboBox cmbCategory;
        private CheckBox chkOnlyRunning;
        private Label lblCount;
        private TextBox txtDetails;
        private Label lblStatus;

        private Button btnToggle;
        private Button btnBatchDisable;
        private Button btnBatchEnable;
        private Button btnSafeOptimize;
        private Button btnBackup;
        private Button btnRestore;
        private Button btnViewLog;
        private Button btnRefresh;
        private ContextMenuStrip ctxMenu;

        // 预设服务分类知识库
        private static readonly Dictionary<string, ServiceMeta> KnowledgeBase = new Dictionary<string, ServiceMeta>(StringComparer.OrdinalIgnoreCase)
        {
            // 1. 推荐安全禁用
            ["DPS"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "诊断策略服务：占用内存与磁盘I/O，禁用仅影响系统自带疑难解答向导。", BadgeColor = Color.SeaGreen },
            ["DiagTrack"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "Connected User Experiences and Telemetry：微软遥测与诊断数据收集服务，建议禁用。", BadgeColor = Color.SeaGreen },
            ["dmwappushservice"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "WAP推送消息路由服务：配合遥测收集发送信息，安全禁用。", BadgeColor = Color.SeaGreen },
            ["TrkWks"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "分布式链接跟踪：局域网移动文件后修复快捷方式，单机完全无用。", BadgeColor = Color.SeaGreen },
            ["QWAVE"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "音视频流媒体QoS调度：局域网流媒体优先级优化，日常无需开启。", BadgeColor = Color.SeaGreen },
            ["CDPSvc"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "连接设备平台服务：用于跨设备同步、手机连接，不用多设备协同可禁用。", BadgeColor = Color.SeaGreen },
            ["CDPUserSvc"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "连接设备平台用户服务：跨设备体验组件（已适配注册表彻底禁用）。", BadgeColor = Color.SeaGreen },
            ["WpnUserService"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "系统推送通知用户服务：禁用后不再接收右下角UWP应用及系统弹窗推送。", BadgeColor = Color.SeaGreen },
            ["TabletInputService"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "触摸键盘和手写面板：台式机或非触屏笔记本可安全禁用。", BadgeColor = Color.SeaGreen },
            ["Browser"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "Computer Browser：维护局域网老旧SMBv1计算机列表，现代网络已淘汰。", BadgeColor = Color.SeaGreen },
            ["MapsBroker"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "下载的地图管理器：不用Win10/11自带地图应用可禁用。", BadgeColor = Color.SeaGreen },
            ["RetailDemo"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "零售演示服务：商场展示样机用，个人电脑可安全禁用。", BadgeColor = Color.SeaGreen },
            ["Fax"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "传真服务：不使用传统电话线传真功能可禁用。", BadgeColor = Color.SeaGreen },
            ["WerSvc"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "Windows错误报告服务：程序崩溃时收集日志并上传微软，可安全禁用。", BadgeColor = Color.SeaGreen },
            ["RemoteRegistry"] = new ServiceMeta { Category = "1. 推荐禁用", Description = "远程注册表：允许远程电脑修改本机注册表，出于安全与性能建议禁用。", BadgeColor = Color.SeaGreen },

            // 2. 按需禁用
            ["SysMain"] = new ServiceMeta { Category = "2. 按需禁用", Description = "超级预读取(原Superfetch)：机械硬盘建议保留，纯固态硬盘(SSD)若频繁高磁盘占用可尝试禁用。", BadgeColor = Color.DarkOrange },
            ["WSearch"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Windows Search：建立文件索引。若已使用 Everything 替代搜索，可禁用以大幅降低磁盘占用。", BadgeColor = Color.DarkOrange },
            ["Spooler"] = new ServiceMeta { Category = "2. 按需禁用", Description = "打印后台处理程序：不使用实体打印机及虚拟PDF打印功能时可禁用。", BadgeColor = Color.DarkOrange },
            ["PrintNotify"] = new ServiceMeta { Category = "2. 按需禁用", Description = "打印机扩展和通知：不用打印机时可配合 Spooler 一同禁用。", BadgeColor = Color.DarkOrange },
            ["BthAvctpSvc"] = new ServiceMeta { Category = "2. 按需禁用", Description = "AVCTP蓝牙音视频控制：不使用蓝牙耳机或蓝牙音箱时可禁用。", BadgeColor = Color.DarkOrange },
            ["bthserv"] = new ServiceMeta { Category = "2. 按需禁用", Description = "蓝牙支持服务：台式机无蓝牙或从不用蓝牙设备时可禁用。", BadgeColor = Color.DarkOrange },
            ["WbioSrvc"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Windows生物识别服务：无指纹识别、无Windows Hello人脸解锁可禁用。", BadgeColor = Color.DarkOrange },
            ["WlanSvc"] = new ServiceMeta { Category = "2. 按需禁用", Description = "WLAN无线配置服务：纯插网线的台式机且永远不用Wi-Fi时可禁用。", BadgeColor = Color.DarkOrange },
            ["LanmanServer"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Server文件共享服务：不需要让局域网其他电脑访问本机共享文件夹时可禁用。", BadgeColor = Color.DarkOrange },
            ["lmhosts"] = new ServiceMeta { Category = "2. 按需禁用", Description = "TCP/IP NetBIOS Helper：辅助局域网共享与名称解析，不用局域网共享可禁用。", BadgeColor = Color.DarkOrange },
            ["RasMan"] = new ServiceMeta { Category = "2. 按需禁用", Description = "远程访问连接管理器：不用系统自带VPN客户端或宽带拨号时可禁用。", BadgeColor = Color.DarkOrange },
            ["SstpSvc"] = new ServiceMeta { Category = "2. 按需禁用", Description = "SSTP VPN协议服务：不使用系统自带SSTP VPN时可禁用（不影响第三方代理）。", BadgeColor = Color.DarkOrange },
            ["camsvc"] = new ServiceMeta { Category = "2. 按需禁用", Description = "功能访问管理器：控制UWP应用对相机/麦克风的权限，不用UWP应用可考虑禁用。", BadgeColor = Color.DarkOrange },
            ["TokenBroker"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Web帐户管理器：用于微软商店与UWP应用账号登录，纯本地账号不使用商店可禁用。", BadgeColor = Color.DarkOrange },
            ["XblAuthManager"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Xbox Live 身份验证管理器：不玩微软商店/Xbox游戏可禁用。", BadgeColor = Color.DarkOrange },
            ["XblGameSave"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Xbox Live 游戏保存：不玩Xbox游戏可禁用。", BadgeColor = Color.DarkOrange },
            ["XboxNetApiSvc"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Xbox Live 网络服务：不玩Xbox联机游戏可禁用。", BadgeColor = Color.DarkOrange },
            ["XboxGipSvc"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Xbox配件管理服务：不使用Xbox无线手柄可禁用。", BadgeColor = Color.DarkOrange },
            ["wuauserv"] = new ServiceMeta { Category = "2. 按需禁用", Description = "Windows Update：控制自动更新，禁用后将无法自动更新系统及安装商店应用。", BadgeColor = Color.DarkOrange },

            // 3. 建议保留/核心系统
            ["AudioEndpointBuilder"] = new ServiceMeta { Category = "3. 核心勿动", Description = "音频端点生成器：禁用会导致电脑完全无声音。", BadgeColor = Color.Firebrick },
            ["Audiosrv"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Windows音频服务：声卡核心服务，禁用会导致无声音。", BadgeColor = Color.Firebrick },
            ["BFE"] = new ServiceMeta { Category = "3. 核心勿动", Description = "基本筛选引擎：防火墙及网络过滤基石，禁用导致网络异常。", BadgeColor = Color.Firebrick },
            ["mpssvc"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Windows Defender 防火墙：禁用会导致部分系统功能及商店异常。", BadgeColor = Color.Firebrick },
            ["WinDefend"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Microsoft Defender 防病毒服务：系统自带防护核心（受PPL保护）。", BadgeColor = Color.Firebrick },
            ["SecurityHealthService"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Windows 安全中心服务：受系统保护。", BadgeColor = Color.Firebrick },
            ["Dhcp"] = new ServiceMeta { Category = "3. 核心勿动", Description = "DHCP客户端：自动获取IP地址，禁用会导致无法上网。", BadgeColor = Color.Firebrick },
            ["Dnscache"] = new ServiceMeta { Category = "3. 核心勿动", Description = "DNS客户端：域名解析缓存，禁用会导致网页无法打开。", BadgeColor = Color.Firebrick },
            ["nsi"] = new ServiceMeta { Category = "3. 核心勿动", Description = "网络存储接口服务：核心网络组件，禁用直接断网。", BadgeColor = Color.Firebrick },
            ["NlaSvc"] = new ServiceMeta { Category = "3. 核心勿动", Description = "网络位置感知：检测网络状态，禁用会导致任务栏网络图标异常。", BadgeColor = Color.Firebrick },
            ["netprofm"] = new ServiceMeta { Category = "3. 核心勿动", Description = "网络列表服务：识别已连接的网络，禁用会导致网络配置异常。", BadgeColor = Color.Firebrick },
            ["Wcmsvc"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Windows连接管理器：管理网络切换与连接状态。", BadgeColor = Color.Firebrick },
            ["LanmanWorkstation"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Workstation客户端：用于访问局域网共享及部分网络映射，建议保留。", BadgeColor = Color.Firebrick },
            ["hidserv"] = new ServiceMeta { Category = "3. 核心勿动", Description = "人体学接口设备服务：支持键盘多媒体快捷键、鼠标侧键等。", BadgeColor = Color.Firebrick },
            ["ShellHWDetection"] = new ServiceMeta { Category = "3. 核心勿动", Description = "外壳硬件检测：U盘、移动硬盘插入自动识别与通知。", BadgeColor = Color.Firebrick },
            ["Schedule"] = new ServiceMeta { Category = "3. 核心勿动", Description = "任务计划程序：系统维护与定时任务核心，禁用会导致多种故障。", BadgeColor = Color.Firebrick },
            ["Winmgmt"] = new ServiceMeta { Category = "3. 核心勿动", Description = "WMI系统管理规范：系统信息查询与驱动管理工具依赖。", BadgeColor = Color.Firebrick },
            ["EventLog"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Windows事件日志：系统日志记录，许多服务依赖它启动。", BadgeColor = Color.Firebrick },
            ["CryptSvc"] = new ServiceMeta { Category = "3. 核心勿动", Description = "加密服务：验证驱动签名与系统证书，禁用会导致安装软件报错。", BadgeColor = Color.Firebrick },
            ["SepMasterService"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Symantec杀毒主服务：受杀毒软件自我保护锁定，无法在此修改。", BadgeColor = Color.Firebrick },
            ["SepScanService"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Symantec扫描服务：受杀毒软件自我保护锁定。", BadgeColor = Color.Firebrick },
            ["sepWscSvc"] = new ServiceMeta { Category = "3. 核心勿动", Description = "Symantec安全中心服务：受杀毒软件自我保护锁定。", BadgeColor = Color.Firebrick },
            ["RpcSs"] = new ServiceMeta { Category = "3. 核心勿动", Description = "远程过程调用(RPC)：Windows最底层通信核心，禁用即崩溃。", BadgeColor = Color.Firebrick },
            ["DcomLaunch"] = new ServiceMeta { Category = "3. 核心勿动", Description = "DCOM服务进程启动器：底层核心，禁用无法开机。", BadgeColor = Color.Firebrick },
            ["LSM"] = new ServiceMeta { Category = "3. 核心勿动", Description = "本地会话管理器：管理用户登录会话，禁用黑屏。", BadgeColor = Color.Firebrick },
            ["SamSs"] = new ServiceMeta { Category = "3. 核心勿动", Description = "安全账户管理器：管理本地账号密码，禁用无法登录。", BadgeColor = Color.Firebrick },
            ["ProfSvc"] = new ServiceMeta { Category = "3. 核心勿动", Description = "用户配置文件服务：加载桌面与个人设置，禁用无法进入桌面。", BadgeColor = Color.Firebrick },
            ["UserManager"] = new ServiceMeta { Category = "3. 核心勿动", Description = "用户管理器：多用户与会话管理核心。", BadgeColor = Color.Firebrick },
            ["Power"] = new ServiceMeta { Category = "3. 核心勿动", Description = "电源管理服务：控制睡眠、关机与CPU电源策略。", BadgeColor = Color.Firebrick },
            ["PlugPlay"] = new ServiceMeta { Category = "3. 核心勿动", Description = "即插即用服务：硬件设备识别核心。", BadgeColor = Color.Firebrick },
            ["StateRepository"] = new ServiceMeta { Category = "3. 核心勿动", Description = "状态存储库服务：开始菜单与现代应用模型依赖，禁用导致开始菜单卡死。", BadgeColor = Color.Firebrick }
        };

        public MainForm()
        {
            _initialBackupPath = Path.Combine(_baseDir, "WinServiceOptimizer_初始备份.json");
            _logFilePath = Path.Combine(_baseDir, "WinServiceOptimizer_操作日志.log");

            Text = "Windows 服务精简优化工具 v2.0 (绿色便携版 · 管理员模式)";
            Size = new Size(1240, 760);
            MinimumSize = new Size(1020, 620);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Microsoft YaHei UI", 9.5f);

            BuildUI();
            BuildContextMenu();
            InitializeDataAsync();
        }

        private void BuildUI()
        {
            // 1. 顶部操作栏
            var actionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 44,
                Padding = new Padding(8, 6, 8, 4),
                WrapContents = false
            };

            btnToggle = MakeButton("切换选中服务", 125, Color.RoyalBlue, Color.White, async (s, e) => await ToggleSelectedAsync());
            btnBatchDisable = MakeButton("禁用勾选", 92, SystemColors.Control, SystemColors.ControlText, async (s, e) => await BatchChangeCheckedAsync(true));
            btnBatchEnable = MakeButton("恢复勾选", 92, SystemColors.Control, SystemColors.ControlText, async (s, e) => await BatchChangeCheckedAsync(false));
            btnSafeOptimize = MakeButton("一键安全精简(推荐项)", 175, Color.SeaGreen, Color.White, async (s, e) => await OneClickSafeOptimizeAsync());
            btnBackup = MakeButton("立即备份", 90, SystemColors.Control, SystemColors.ControlText, (s, e) => CreateManualBackup());
            btnRestore = MakeButton("还原快照", 90, SystemColors.Control, SystemColors.ControlText, async (s, e) => await RestoreBackupAsync());
            btnViewLog = MakeButton("查看日志", 90, SystemColors.Control, SystemColors.ControlText, (s, e) => ShowLogDialog());
            btnRefresh = MakeButton("刷新列表", 90, SystemColors.Control, SystemColors.ControlText, async (s, e) => await ReloadServicesAsync());

            actionPanel.Controls.AddRange(new Control[]
            {
                btnToggle, btnBatchDisable, btnBatchEnable, btnSafeOptimize,
                btnBackup, btnRestore, btnViewLog, btnRefresh
            });

            // 2. 顶部搜索与筛选栏
            var filterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 38,
                Padding = new Padding(10, 4, 10, 4),
                WrapContents = false
            };

            var lblSearch = new Label { Text = "搜索:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) };
            txtSearch = new TextBox { Width = 230, Margin = new Padding(0, 2, 14, 0) };
            txtSearch.TextChanged += (s, e) => ApplyFilterAndRender();

            var lblCat = new Label { Text = "筛选:", AutoSize = true, Margin = new Padding(0, 6, 4, 0) };
            cmbCategory = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 175,
                Margin = new Padding(0, 2, 14, 0)
            };
            cmbCategory.Items.AddRange(new object[]
            {
                "全部分类",
                "1. 推荐禁用",
                "2. 按需禁用",
                "3. 核心勿动",
                "4. 其他系统服务",
                "仅看第三方软件服务",
                "仅看已禁用服务"
            });
            cmbCategory.SelectedIndex = 0;
            cmbCategory.SelectedIndexChanged += (s, e) => ApplyFilterAndRender();

            chkOnlyRunning = new CheckBox
            {
                Text = "仅显示【正在运行】的服务",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(0, 5, 16, 0)
            };
            chkOnlyRunning.CheckedChanged += (s, e) => ApplyFilterAndRender();

            lblCount = new Label
            {
                Text = "加载中...",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Margin = new Padding(0, 6, 0, 0)
            };

            filterPanel.Controls.AddRange(new Control[]
            {
                lblSearch, txtSearch, lblCat, cmbCategory, chkOnlyRunning, lblCount
            });

            // 3. 底部状态栏
            var statusPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Padding = new Padding(10, 2, 10, 2),
                BackColor = SystemColors.ControlLight
            };
            lblStatus = new Label
            {
                Text = "提示：绿色推荐禁用，橙色按需禁用，红色核心勿动；所有日志与备份均保存在当前程序目录。",
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(60, 60, 60),
                TextAlign = ContentAlignment.MiddleLeft
            };
            statusPanel.Controls.Add(lblStatus);

            // 4. 底部详细信息面板
            var detailPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 148,
                Padding = new Padding(10, 4, 10, 4)
            };
            var grpDetail = new GroupBox
            {
                Text = "选中服务详细信息（完整描述 / 程序路径 / 依赖关系）",
                Dock = DockStyle.Fill,
                Padding = new Padding(6, 4, 6, 6)
            };
            txtDetails = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Window,
                Text = "点击列表中任意服务即可在此查看完整信息..."
            };
            grpDetail.Controls.Add(txtDetails);
            detailPanel.Controls.Add(grpDetail);

            // 5. 中间主列表
            listView = new BufferedListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                CheckBoxes = true,
                MultiSelect = false,
                ListViewItemSorter = _sorter
            };

            listView.Columns.Add("优化分类", 110);
            listView.Columns.Add("服务名称", 150);
            listView.Columns.Add("运行状态", 78);
            listView.Columns.Add("启动类型", 96);
            listView.Columns.Add("来源厂商", 125);
            listView.Columns.Add("内存占用", 85);
            listView.Columns.Add("显示名称", 200);
            listView.Columns.Add("功能与禁用影响说明", 350);

            listView.ColumnClick += ListView_ColumnClick;
            listView.SelectedIndexChanged += ListView_SelectedIndexChanged;
            listView.DoubleClick += async (s, e) => await ToggleSelectedAsync();

            // 按 Docking 从内到外添加顺序
            Controls.Add(listView);
            Controls.Add(detailPanel);
            Controls.Add(statusPanel);
            Controls.Add(filterPanel);
            Controls.Add(actionPanel);
        }

        private Button MakeButton(string text, int width, Color backColor, Color foreColor, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Width = width,
                Height = 30,
                Margin = new Padding(0, 0, 8, 0),
                UseVisualStyleBackColor = backColor == SystemColors.Control
            };
            if (backColor != SystemColors.Control)
            {
                btn.BackColor = backColor;
                btn.ForeColor = foreColor;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
            }
            btn.Click += onClick;
            return btn;
        }

        private void BuildContextMenu()
        {
            ctxMenu = new ContextMenuStrip();

            var miStart = new ToolStripMenuItem("立即启动服务", null, async (s, e) => await ControlSelectedRunStateAsync("start"));
            var miStop = new ToolStripMenuItem("立即停止服务", null, async (s, e) => await ControlSelectedRunStateAsync("stop"));
            var miRestart = new ToolStripMenuItem("重启服务", null, async (s, e) => await ControlSelectedRunStateAsync("restart"));

            var miSetAuto = new ToolStripMenuItem("设为：自动启动", null, async (s, e) => await SetSelectedStartModeAsync(2, false));
            var miSetDelayed = new ToolStripMenuItem("设为：自动 (延迟启动)", null, async (s, e) => await SetSelectedStartModeAsync(2, true));
            var miSetManual = new ToolStripMenuItem("设为：手动启动", null, async (s, e) => await SetSelectedStartModeAsync(3, false));
            var miSetDisabled = new ToolStripMenuItem("设为：已禁用", null, async (s, e) => await SetSelectedStartModeAsync(4, false));

            var miOpenDir = new ToolStripMenuItem("打开程序所在文件夹", null, (s, e) => OpenSelectedExeFolder());
            var miCopyName = new ToolStripMenuItem("复制服务名称", null, (s, e) => CopySelectedServiceName());
            var miOpenMsc = new ToolStripMenuItem("打开系统服务管理器 (services.msc)", null, (s, e) =>
            {
                try { Process.Start(new ProcessStartInfo("services.msc") { UseShellExecute = true }); }
                catch (Exception ex) { MessageBox.Show("无法打开 services.msc：" + ex.Message); }
            });

            ctxMenu.Items.AddRange(new ToolStripItem[]
            {
                miStart, miStop, miRestart,
                new ToolStripSeparator(),
                miSetAuto, miSetDelayed, miSetManual, miSetDisabled,
                new ToolStripSeparator(),
                miOpenDir, miCopyName, miOpenMsc
            });

            listView.ContextMenuStrip = ctxMenu;
        }

        private async void InitializeDataAsync()
        {
            await ReloadServicesAsync();
            EnsureInitialBackup();
        }

        private void EnsureInitialBackup()
        {
            if (File.Exists(_initialBackupPath))
            {
                var loaded = LoadBackupFile(_initialBackupPath);
                if (loaded != null)
                {
                    _initialSnapshotMap = loaded.Services.ToDictionary(s => s.ServiceName, StringComparer.OrdinalIgnoreCase);
                    return;
                }
            }

            var snapshot = BuildSnapshotFromCurrentList();
            if (SaveBackupFile(_initialBackupPath, snapshot))
            {
                _initialSnapshotMap = snapshot.Services.ToDictionary(s => s.ServiceName, StringComparer.OrdinalIgnoreCase);
                WriteLog("首次运行，已自动生成初始备份快照：" + Path.GetFileName(_initialBackupPath));
                lblStatus.Text = "已自动在当前目录生成初始备份快照 (WinServiceOptimizer_初始备份.json)。";
            }
        }

        private void SetBusy(bool busy, string statusText = "")
        {
            _isBusy = busy;
            btnToggle.Enabled = !busy;
            btnBatchDisable.Enabled = !busy;
            btnBatchEnable.Enabled = !busy;
            btnSafeOptimize.Enabled = !busy;
            btnBackup.Enabled = !busy;
            btnRestore.Enabled = !busy;
            btnRefresh.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
            if (!string.IsNullOrEmpty(statusText))
                lblStatus.Text = statusText;
        }

        private async Task ReloadServicesAsync()
        {
            if (_isBusy) return;
            SetBusy(true, "正在扫描系统服务、进程内存与注册表信息...");

            try
            {
                var data = await Task.Run(() => ScanAllServices());
                _allServices = data;
                ApplyFilterAndRender();
                lblStatus.Text = $"扫描完成，共加载 {_allServices.Count} 项系统服务。";
            }
            catch (Exception ex)
            {
                MessageBox.Show("扫描服务失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private List<ServiceRowInfo> ScanAllServices()
        {
            var result = new List<ServiceRowInfo>();
            var pidMemoryCache = new Dictionary<int, double>();
            var pidServiceCount = new Dictionary<int, int>();
            var tempEntries = new List<(ServiceRowInfo info, int pid)>();

            IntPtr hScm = OpenSCManager(null, null, SC_MANAGER_CONNECT);
            try
            {
                var services = ServiceController.GetServices();
                foreach (var sc in services)
                {
                    using (sc)
                    {
                        var meta = GetMeta(sc.ServiceName);
                        bool isRunning = sc.Status == ServiceControllerStatus.Running;

                        int startVal = 3;
                        bool delayedAuto = false;
                        string rawImagePath = "";
                        string rawDesc = "";

                        try
                        {
                            using var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{sc.ServiceName}");
                            if (key != null)
                            {
                                if (key.GetValue("Start") is int sv) startVal = sv;
                                if (key.GetValue("DelayedAutostart") is int da && da == 1) delayedAuto = true;
                                rawImagePath = key.GetValue("ImagePath")?.ToString() ?? "";
                                rawDesc = key.GetValue("Description")?.ToString() ?? "";
                            }
                        }
                        catch { }

                        string startText = FormatStartModeText(startVal, delayedAuto);
                        string exePath = ExtractExecutablePath(rawImagePath);
                        (string sourceLabel, bool isThirdParty) = DetectServiceSource(exePath);
                        string officialDesc = ResolveDescriptionString(rawDesc);

                        var dependsOn = new List<string>();
                        var dependedBy = new List<string>();
                        try
                        {
                            foreach (var dep in sc.ServicesDependedOn)
                            {
                                using (dep) { dependsOn.Add($"{dep.DisplayName} ({dep.ServiceName})"); }
                            }
                            foreach (var dep in sc.DependentServices)
                            {
                                using (dep) { dependedBy.Add($"{dep.DisplayName} ({dep.ServiceName})"); }
                            }
                        }
                        catch { }

                        int pid = 0;
                        if (isRunning && hScm != IntPtr.Zero)
                        {
                            pid = QueryServicePid(hScm, sc.ServiceName);
                            if (pid > 0)
                            {
                                pidServiceCount[pid] = pidServiceCount.TryGetValue(pid, out int c) ? c + 1 : 1;
                            }
                        }

                        string advice = meta.Description;
                        if (meta.Category.StartsWith("4.") && !string.IsNullOrEmpty(officialDesc))
                        {
                            advice = officialDesc;
                        }

                        var info = new ServiceRowInfo
                        {
                            Category = meta.Category,
                            BadgeColor = meta.BadgeColor,
                            ServiceName = sc.ServiceName,
                            DisplayName = sc.DisplayName,
                            IsRunning = isRunning,
                            StartModeValue = startVal,
                            DelayedAutoStart = delayedAuto,
                            StartModeText = startText,
                            SourceLabel = sourceLabel,
                            IsThirdParty = isThirdParty,
                            ExePath = exePath,
                            RawImagePath = rawImagePath,
                            ProcessId = pid,
                            OfficialDescription = officialDesc,
                            AdviceDescription = advice,
                            DependsOn = dependsOn,
                            DependedBy = dependedBy
                        };

                        tempEntries.Add((info, pid));
                    }
                }
            }
            finally
            {
                if (hScm != IntPtr.Zero) CloseServiceHandle(hScm);
            }

            foreach (var entry in tempEntries)
            {
                var info = entry.info;
                int pid = entry.pid;

                if (pid > 0)
                {
                    if (!pidMemoryCache.TryGetValue(pid, out double memMb))
                    {
                        memMb = GetProcessMemoryMb(pid);
                        pidMemoryCache[pid] = memMb;
                    }
                    info.MemoryMb = memMb;
                    info.IsSharedHost = pidServiceCount.TryGetValue(pid, out int cnt) && cnt > 1;
                    if (memMb > 0)
                    {
                        info.MemoryText = info.IsSharedHost ? $"{memMb:F1} MB*" : $"{memMb:F1} MB";
                    }
                }

                result.Add(info);
            }

            return result;
        }

        private void ApplyFilterAndRender()
        {
            string keyword = txtSearch.Text.Trim();
            string catFilter = cmbCategory.SelectedItem?.ToString() ?? "全部分类";
            bool onlyRunning = chkOnlyRunning.Checked;

            listView.BeginUpdate();
            listView.Items.Clear();

            int shown = 0;
            foreach (var info in _allServices)
            {
                if (onlyRunning && !info.IsRunning)
                    continue;

                if (catFilter == "仅看第三方软件服务")
                {
                    if (!info.IsThirdParty) continue;
                }
                else if (catFilter == "仅看已禁用服务")
                {
                    if (info.StartModeValue != 4) continue;
                }
                else if (catFilter != "全部分类")
                {
                    if (!info.Category.Equals(catFilter, StringComparison.OrdinalIgnoreCase)) continue;
                }

                if (!string.IsNullOrEmpty(keyword))
                {
                    bool hit =
                        info.ServiceName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        info.DisplayName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        info.AdviceDescription.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        info.OfficialDescription.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        info.SourceLabel.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        info.ExePath.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;

                    if (!hit) continue;
                }

                var lvi = new ListViewItem(info.Category)
                {
                    ForeColor = info.BadgeColor,
                    Tag = info
                };
                lvi.SubItems.Add(info.ServiceName);
                lvi.SubItems.Add(info.IsRunning ? "正在运行" : "已停止");
                lvi.SubItems.Add(info.StartModeText);
                lvi.SubItems.Add(info.SourceLabel);
                lvi.SubItems.Add(info.MemoryText);
                lvi.SubItems.Add(info.DisplayName);
                lvi.SubItems.Add(info.AdviceDescription);

                listView.Items.Add(lvi);
                shown++;
            }

            listView.Sort();
            listView.EndUpdate();

            lblCount.Text = $"显示: {shown} / 总计: {_allServices.Count} 项";
        }

        private void ListView_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (e.Column == _sorter.SortColumn)
            {
                _sorter.Order = _sorter.Order == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                _sorter.SortColumn = e.Column;
                _sorter.Order = SortOrder.Ascending;
            }
            listView.Sort();
        }

        private void ListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
                return;

            if (!(listView.SelectedItems[0].Tag is ServiceRowInfo info))
                return;

            var sb = new StringBuilder();
            sb.AppendLine($"【服务名称】{info.ServiceName}    【显示名称】{info.DisplayName}");
            sb.AppendLine($"【优化分类】{info.Category}    【来源】{info.SourceLabel}    【状态】{(info.IsRunning ? "正在运行" : "已停止")} ({info.StartModeText})");
            if (info.ProcessId > 0)
            {
                string hostTag = info.IsSharedHost ? "（与其他服务共享宿主进程，标 * 为总进程内存）" : "（独立进程）";
                sb.AppendLine($"【进程信息】PID: {info.ProcessId}    占用内存: {info.MemoryMb:F1} MB {hostTag}");
            }
            sb.AppendLine($"【程序路径】{(string.IsNullOrEmpty(info.RawImagePath) ? "未配置" : info.RawImagePath)}");
            sb.AppendLine($"【建议说明】{info.AdviceDescription}");
            if (!string.IsNullOrEmpty(info.OfficialDescription) && info.OfficialDescription != info.AdviceDescription)
                sb.AppendLine($"【官方描述】{info.OfficialDescription}");
            if (info.DependsOn.Count > 0)
                sb.AppendLine($"【它依赖的服务】{string.Join("；", info.DependsOn)}");
            if (info.DependedBy.Count > 0)
                sb.AppendLine($"【依赖它的服务】{string.Join("；", info.DependedBy)}");

            txtDetails.Text = sb.ToString();
        }

        private async Task ToggleSelectedAsync()
        {
            if (_isBusy) return;
            if (listView.SelectedItems.Count == 0)
            {
                MessageBox.Show("请先点击选择一个服务！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var info = (ServiceRowInfo)listView.SelectedItems[0].Tag;
            bool toDisable = info.StartModeValue != 4;

            if (toDisable)
            {
                if (!ConfirmRiskyAndDependencies(info, stopOrDisable: true))
                    return;

                await ApplyServiceChangesAsync(new List<(ServiceRowInfo info, int targetStart, bool delayed, bool tryStart, bool tryStop)>
                {
                    (info, 4, false, false, true)
                }, "禁用服务");
            }
            else
            {
                // 恢复启用：优先恢复成初始备份时的原始启动类型
                int targetStart = 3;
                bool delayed = false;
                if (_initialSnapshotMap.TryGetValue(info.ServiceName, out var orig) && orig.StartMode != 4)
                {
                    targetStart = orig.StartMode;
                    delayed = orig.DelayedAutostart == 1;
                }

                await ApplyServiceChangesAsync(new List<(ServiceRowInfo info, int targetStart, bool delayed, bool tryStart, bool tryStop)>
                {
                    (info, targetStart, delayed, true, false)
                }, "恢复启用");
            }
        }

        private async Task BatchChangeCheckedAsync(bool disable)
        {
            if (_isBusy) return;
            if (listView.CheckedItems.Count == 0)
            {
                MessageBox.Show("请先在列表左侧勾选至少一个服务！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selectedInfos = listView.CheckedItems.Cast<ListViewItem>()
                .Select(i => i.Tag as ServiceRowInfo)
                .Where(i => i != null)
                .ToList();

            if (disable && selectedInfos.Any(i => i.Category.StartsWith("3.")))
            {
                var coreNames = string.Join("、", selectedInfos.Where(i => i.Category.StartsWith("3.")).Select(i => i.ServiceName));
                var warn = MessageBox.Show(
                    $"你勾选的列表中包含【系统核心/杀毒保护】服务：\n{coreNames}\n\n强行禁用可能导致无法开机或断网，确定仍要继续吗？",
                    "高危批量操作警告", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
                if (warn != DialogResult.Yes) return;
            }
            else
            {
                var confirm = MessageBox.Show(
                    $"确定要批量{(disable ? "禁用并尝试停止" : "恢复启用并启动")}所勾选的 {selectedInfos.Count} 项服务吗？",
                    "确认批量操作", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirm != DialogResult.Yes) return;
            }

            var tasks = new List<(ServiceRowInfo info, int targetStart, bool delayed, bool tryStart, bool tryStop)>();
            foreach (var info in selectedInfos)
            {
                if (disable)
                {
                    tasks.Add((info, 4, false, false, true));
                }
                else
                {
                    int targetStart = 3;
                    bool delayed = false;
                    if (_initialSnapshotMap.TryGetValue(info.ServiceName, out var orig) && orig.StartMode != 4)
                    {
                        targetStart = orig.StartMode;
                        delayed = orig.DelayedAutostart == 1;
                    }
                    tasks.Add((info, targetStart, delayed, true, false));
                }
            }

            await ApplyServiceChangesAsync(tasks, disable ? "批量禁用" : "批量恢复");
        }

        private async Task OneClickSafeOptimizeAsync()
        {
            if (_isBusy) return;

            var candidates = _allServices
                .Where(s => s.Category.StartsWith("1.") && (s.StartModeValue != 4 || s.IsRunning))
                .ToList();

            if (candidates.Count == 0)
            {
                MessageBox.Show("当前所有【1. 推荐禁用】的服务均已处于禁用停止状态，无需再次优化！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var namesPreview = string.Join("\n", candidates.Select(c => $"• {c.ServiceName} ({c.DisplayName})"));
            var confirm = MessageBox.Show(
                $"即将安全禁用并停止以下 {candidates.Count} 项【绿色推荐禁用】服务（随时可通过“还原快照”恢复）：\n\n{namesPreview}\n\n确定立即执行一键优化吗？",
                "一键安全精简确认", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (confirm != DialogResult.Yes) return;

            var tasks = candidates.Select(c => (c, 4, false, false, true)).ToList();
            await ApplyServiceChangesAsync(tasks, "一键安全精简");
        }

        private async Task SetSelectedStartModeAsync(int targetStart, bool delayed)
        {
            if (_isBusy || listView.SelectedItems.Count == 0) return;
            var info = (ServiceRowInfo)listView.SelectedItems[0].Tag;

            if (targetStart == 4 && !ConfirmRiskyAndDependencies(info, stopOrDisable: false))
                return;

            await ApplyServiceChangesAsync(new List<(ServiceRowInfo info, int targetStart, bool delayed, bool tryStart, bool tryStop)>
            {
                (info, targetStart, delayed, false, false)
            }, "修改启动类型");
        }

        private async Task ControlSelectedRunStateAsync(string action)
        {
            if (_isBusy || listView.SelectedItems.Count == 0) return;
            var info = (ServiceRowInfo)listView.SelectedItems[0].Tag;

            if ((action == "stop" || action == "restart") && !ConfirmRiskyAndDependencies(info, stopOrDisable: true))
                return;

            SetBusy(true, $"正在对 [{info.ServiceName}] 执行 {action} 操作...");
            string resultMsg = "";
            bool hasError = false;

            await Task.Run(() =>
            {
                try
                {
                    using var sc = new ServiceController(info.ServiceName);
                    if (action == "start")
                    {
                        if (info.StartModeValue == 4)
                            throw new InvalidOperationException("该服务当前处于【已禁用】状态，请先修改启动类型再启动。");
                        if (sc.Status != ServiceControllerStatus.Running)
                        {
                            sc.Start();
                            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(8));
                        }
                        resultMsg = $"已成功启动 [{info.ServiceName}]";
                    }
                    else if (action == "stop")
                    {
                        if (sc.Status != ServiceControllerStatus.Stopped)
                        {
                            if (!sc.CanStop)
                                throw new InvalidOperationException("该服务拒绝即时停止请求（CanStop = false）。");
                            sc.Stop();
                            sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(8));
                        }
                        resultMsg = $"已成功停止 [{info.ServiceName}]";
                    }
                    else if (action == "restart")
                    {
                        if (sc.Status != ServiceControllerStatus.Stopped)
                        {
                            if (!sc.CanStop)
                                throw new InvalidOperationException("该服务拒绝即时停止请求，无法重启。");
                            sc.Stop();
                            sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(8));
                        }
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(8));
                        resultMsg = $"已成功重启 [{info.ServiceName}]";
                    }
                }
                catch (Exception ex)
                {
                    hasError = true;
                    resultMsg = $"对 [{info.ServiceName}] 执行 {action} 失败：{ex.Message}";
                }
            });

            WriteLog(resultMsg);
            SetBusy(false);

            if (hasError)
                MessageBox.Show(resultMsg, "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);

            await ReloadServicesAsync();
            lblStatus.Text = resultMsg;
        }

        private bool ConfirmRiskyAndDependencies(ServiceRowInfo info, bool stopOrDisable)
        {
            if (info.Category.StartsWith("3."))
            {
                var warn = MessageBox.Show(
                    $"【{info.ServiceName}】属于系统核心或杀毒软件保护服务，强行操作可能导致无法开机、断网或系统崩溃！\n\n确定仍要继续吗？",
                    "高危操作警告", MessageBoxButtons.YesNo, MessageBoxIcon.Stop);
                if (warn != DialogResult.Yes) return false;
            }

            if (stopOrDisable && info.IsRunning)
            {
                try
                {
                    using var sc = new ServiceController(info.ServiceName);
                    var activeDeps = new List<string>();
                    foreach (var dep in sc.DependentServices)
                    {
                        using (dep)
                        {
                            if (dep.Status != ServiceControllerStatus.Stopped)
                                activeDeps.Add($"• {dep.DisplayName} ({dep.ServiceName})");
                        }
                    }

                    if (activeDeps.Count > 0)
                    {
                        var depWarn = MessageBox.Show(
                            $"注意：当前有 {activeDeps.Count} 个正在运行的服务依赖 [{info.ServiceName}]：\n\n{string.Join("\n", activeDeps)}\n\n停止它可能会导致上述服务异常或被联动停止，确定继续吗？",
                            "服务依赖关系提醒", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (depWarn != DialogResult.Yes) return false;
                    }
                }
                catch { }
            }

            return true;
        }

        private async Task ApplyServiceChangesAsync(
            List<(ServiceRowInfo info, int targetStart, bool delayed, bool tryStart, bool tryStop)> items,
            string actionTitle)
        {
            SetBusy(true, $"正在执行 {actionTitle}（共 {items.Count} 项）...");

            var summaryLogs = new List<string>();
            int successCount = 0;
            int needRebootCount = 0;
            int failCount = 0;

            await Task.Run(() =>
            {
                foreach (var item in items)
                {
                    var info = item.info;
                    string oldMode = info.StartModeText;
                    string newMode = FormatStartModeText(item.targetStart, item.delayed);

                    try
                    {
                        SetServiceStartConfig(info.ServiceName, item.targetStart, item.delayed);

                        string runtimeNote = "";
                        using (var sc = new ServiceController(info.ServiceName))
                        {
                            if (item.tryStop && sc.Status != ServiceControllerStatus.Stopped)
                            {
                                if (sc.CanStop)
                                {
                                    try
                                    {
                                        sc.Stop();
                                        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(6));
                                        runtimeNote = "并已停止";
                                    }
                                    catch
                                    {
                                        needRebootCount++;
                                        runtimeNote = "（热停止超时/受限，需重启电脑生效）";
                                    }
                                }
                                else
                                {
                                    needRebootCount++;
                                    runtimeNote = "（当前不支持热停止，需重启电脑生效）";
                                }
                            }
                            else if (item.tryStart && sc.Status == ServiceControllerStatus.Stopped && item.targetStart != 4)
                            {
                                try
                                {
                                    sc.Start();
                                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(6));
                                    runtimeNote = "并已启动";
                                }
                                catch (Exception startEx)
                                {
                                    runtimeNote = $"（配置已改，但启动未成功：{startEx.Message}）";
                                }
                            }
                        }

                        successCount++;
                        string msg = $"[{actionTitle}] {info.ServiceName} ({info.DisplayName}): {oldMode} -> {newMode} {runtimeNote}".Trim();
                        summaryLogs.Add(msg);
                        WriteLog(msg);
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        string errMsg = $"[{actionTitle}失败] {info.ServiceName}: {ex.Message}";
                        summaryLogs.Add(errMsg);
                        WriteLog(errMsg);
                    }
                }
            });

            SetBusy(false);
            await ReloadServicesAsync();

            string finalStatus = $"{actionTitle}完成：成功 {successCount} 项" +
                                 (needRebootCount > 0 ? $"（其中 {needRebootCount} 项需重启生效）" : "") +
                                 (failCount > 0 ? $"，失败 {failCount} 项" : "") + "。";
            lblStatus.Text = finalStatus;

            if (items.Count == 1 && failCount == 1)
            {
                MessageBox.Show(summaryLogs[0] + "\n\n注：部分受系统PPL或杀毒软件自我保护的服务无法直接更改。", "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (items.Count == 1 && needRebootCount == 1)
            {
                MessageBox.Show(summaryLogs[0], "需要重启生效", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (items.Count > 1)
            {
                MessageBox.Show(finalStatus + "\n\n详细记录已写入当前目录的操作日志。", actionTitle + "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SetServiceStartConfig(string serviceName, int targetStart, bool delayedAutoStart)
        {
            bool scmUpdated = false;
            IntPtr hScm = OpenSCManager(null, null, SC_MANAGER_CONNECT);
            if (hScm != IntPtr.Zero)
            {
                try
                {
                    IntPtr hSvc = OpenService(hScm, serviceName, SERVICE_CHANGE_CONFIG);
                    if (hSvc != IntPtr.Zero)
                    {
                        try
                        {
                            scmUpdated = ChangeServiceConfig(
                                hSvc, SERVICE_NO_CHANGE, (uint)targetStart, SERVICE_NO_CHANGE,
                                null, null, IntPtr.Zero, null, null, null, null);
                        }
                        finally
                        {
                            CloseServiceHandle(hSvc);
                        }
                    }
                }
                finally
                {
                    CloseServiceHandle(hScm);
                }
            }

            // 同步写入注册表（确保 DelayedAutostart 与 Per-User 服务模板一并生效）
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{serviceName}", true);
                if (key == null)
                {
                    if (!scmUpdated) throw new InvalidOperationException("注册表服务键不存在或无权访问。");
                }
                else
                {
                    key.SetValue("Start", targetStart, RegistryValueKind.DWord);
                    if (targetStart == 2 && delayedAutoStart)
                        key.SetValue("DelayedAutostart", 1, RegistryValueKind.DWord);
                    else if (key.GetValue("DelayedAutostart") != null)
                        key.SetValue("DelayedAutostart", 0, RegistryValueKind.DWord);
                }

                // 处理带随机后缀的 Per-User 服务模板（如 CDPUserSvc_xxxxx），防止重启后复原
                int idx = serviceName.LastIndexOf('_');
                if (idx > 0)
                {
                    string baseName = serviceName.Substring(0, idx);
                    using var baseKey = Registry.LocalMachine.OpenSubKey($@"SYSTEM\CurrentControlSet\Services\{baseName}", true);
                    if (baseKey != null)
                    {
                        baseKey.SetValue("Start", targetStart, RegistryValueKind.DWord);
                        if (targetStart == 2 && delayedAutoStart)
                            baseKey.SetValue("DelayedAutostart", 1, RegistryValueKind.DWord);
                        else if (baseKey.GetValue("DelayedAutostart") != null)
                            baseKey.SetValue("DelayedAutostart", 0, RegistryValueKind.DWord);
                    }
                }
            }
            catch (Exception ex)
            {
                if (!scmUpdated)
                    throw new InvalidOperationException("权限被拒绝或受系统保护：" + ex.Message, ex);
            }
        }

        private BackupSnapshot BuildSnapshotFromCurrentList()
        {
            var snap = new BackupSnapshot
            {
                CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                MachineName = Environment.MachineName
            };

            foreach (var s in _allServices)
            {
                snap.Services.Add(new ServiceBackupEntry
                {
                    ServiceName = s.ServiceName,
                    DisplayName = s.DisplayName,
                    StartMode = s.StartModeValue,
                    DelayedAutostart = s.DelayedAutoStart ? 1 : 0,
                    WasRunning = s.IsRunning
                });
            }
            return snap;
        }

        private void CreateManualBackup()
        {
            string fileName = $"WinServiceOptimizer_备份_{DateTime.Now:yyyyMMdd_HHmmss}.json";
            string fullPath = Path.Combine(_baseDir, fileName);
            var snap = BuildSnapshotFromCurrentList();

            if (SaveBackupFile(fullPath, snap))
            {
                WriteLog("手动导出备份快照：" + fileName);
                lblStatus.Text = "已保存备份快照至当前目录：" + fileName;
                MessageBox.Show($"备份成功！已保存至程序当前目录：\n{fullPath}", "备份完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private async Task RestoreBackupAsync()
        {
            if (_isBusy) return;

            using var ofd = new OpenFileDialog
            {
                Title = "选择要还原的备份快照文件",
                InitialDirectory = _baseDir,
                Filter = "服务备份快照 (*.json)|*.json",
                FileName = File.Exists(_initialBackupPath) ? Path.GetFileName(_initialBackupPath) : ""
            };

            if (ofd.ShowDialog() != DialogResult.OK) return;

            var snap = LoadBackupFile(ofd.FileName);
            if (snap == null || snap.Services.Count == 0)
            {
                MessageBox.Show("无法读取该备份文件，或文件格式无效！", "还原失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 只找出当前与备份存在差异的服务，避免乱动核心未改动服务
            var currentMap = _allServices.ToDictionary(s => s.ServiceName, StringComparer.OrdinalIgnoreCase);
            var diffs = new List<(ServiceRowInfo current, ServiceBackupEntry target)>();

            foreach (var entry in snap.Services)
            {
                if (!currentMap.TryGetValue(entry.ServiceName, out var cur))
                    continue;

                bool startChanged = cur.StartModeValue != entry.StartMode;
                bool delayChanged = (cur.DelayedAutoStart ? 1 : 0) != entry.DelayedAutostart;
                if (startChanged || delayChanged)
                {
                    diffs.Add((cur, entry));
                }
            }

            if (diffs.Count == 0)
            {
                MessageBox.Show($"当前所有服务的启动类型与快照 ({snap.CreatedAt}) 完全一致，无需还原！", "无需还原", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var preview = string.Join("\n", diffs.Take(15).Select(d =>
                $"• {d.current.ServiceName}: {d.current.StartModeText} -> {FormatStartModeText(d.target.StartMode, d.target.DelayedAutostart == 1)}"));
            if (diffs.Count > 15) preview += $"\n... 等共 {diffs.Count} 项";

            var confirm = MessageBox.Show(
                $"备份创建时间：{snap.CreatedAt}\n检测到 {diffs.Count} 项服务的启动配置与该快照不同：\n\n{preview}\n\n是否立即将这些服务还原到备份时的状态？",
                "确认还原快照", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirm != DialogResult.Yes) return;

            var tasks = diffs.Select(d => (
                d.current,
                d.target.StartMode,
                d.target.DelayedAutostart == 1,
                d.target.WasRunning && d.target.StartMode != 4,
                !d.target.WasRunning && d.target.StartMode == 4
            )).ToList();

            await ApplyServiceChangesAsync(tasks, "还原快照");
        }

        private bool SaveBackupFile(string path, BackupSnapshot snapshot)
        {
            try
            {
                using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read);
                using var writer = JsonReaderWriterFactory.CreateJsonWriter(fs, Encoding.UTF8, false, true, "  ");
                var ser = new DataContractJsonSerializer(typeof(BackupSnapshot));
                ser.WriteObject(writer, snapshot);
                writer.Flush();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存备份文件失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private BackupSnapshot LoadBackupFile(string path)
        {
            try
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var ser = new DataContractJsonSerializer(typeof(BackupSnapshot));
                return ser.ReadObject(fs) as BackupSnapshot;
            }
            catch
            {
                return null;
            }
        }

        private void WriteLog(string message)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            lock (_sessionLogs)
            {
                _sessionLogs.Add(line);
                try
                {
                    File.AppendAllText(_logFilePath, line + Environment.NewLine, Encoding.UTF8);
                }
                catch { }
            }
        }

        private void ShowLogDialog()
        {
            using var dlg = new Form
            {
                Text = "操作日志记录 (保存位置：" + _logFilePath + ")",
                Size = new Size(780, 480),
                StartPosition = FormStartPosition.CenterParent,
                Font = Font
            };

            var txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false
            };

            if (File.Exists(_logFilePath))
            {
                try { txtLog.Text = File.ReadAllText(_logFilePath, Encoding.UTF8); }
                catch { txtLog.Text = string.Join(Environment.NewLine, _sessionLogs); }
            }
            else
            {
                txtLog.Text = _sessionLogs.Count > 0 ? string.Join(Environment.NewLine, _sessionLogs) : "暂无操作日志记录。";
            }

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 42, Padding = new Padding(8, 6, 8, 6), FlowDirection = FlowDirection.RightToLeft };
            var btnClose = new Button { Text = "关闭", Width = 85, Height = 28 };
            btnClose.Click += (s, e) => dlg.Close();

            var btnOpenFolder = new Button { Text = "打开程序目录", Width = 115, Height = 28 };
            btnOpenFolder.Click += (s, e) =>
            {
                try { Process.Start("explorer.exe", _baseDir); } catch { }
            };

            bottom.Controls.Add(btnClose);
            bottom.Controls.Add(btnOpenFolder);

            dlg.Controls.Add(txtLog);
            dlg.Controls.Add(bottom);
            dlg.ShowDialog(this);
        }

        private void OpenSelectedExeFolder()
        {
            if (listView.SelectedItems.Count == 0) return;
            var info = (ServiceRowInfo)listView.SelectedItems[0].Tag;
            if (string.IsNullOrEmpty(info.ExePath) || !File.Exists(info.ExePath))
            {
                MessageBox.Show($"未能定位到程序文件：\n{info.RawImagePath}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                Process.Start("explorer.exe", $"/select,\"{info.ExePath}\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开文件夹失败：" + ex.Message);
            }
        }

        private void CopySelectedServiceName()
        {
            if (listView.SelectedItems.Count == 0) return;
            var info = (ServiceRowInfo)listView.SelectedItems[0].Tag;
            Clipboard.SetText(info.ServiceName);
            lblStatus.Text = $"已复制服务名称：{info.ServiceName}";
        }

        private static ServiceMeta GetMeta(string serviceName)
        {
            if (KnowledgeBase.TryGetValue(serviceName, out var meta))
                return meta;

            int idx = serviceName.LastIndexOf('_');
            if (idx > 0)
            {
                string prefix = serviceName.Substring(0, idx);
                if (KnowledgeBase.TryGetValue(prefix, out var prefixMeta))
                    return prefixMeta;
            }

            return new ServiceMeta
            {
                Category = "4. 其他系统服务",
                Description = "系统常规或第三方组件服务，若不确定用途建议查看下方官方描述与程序路径。",
                BadgeColor = Color.DimGray
            };
        }

        private static string FormatStartModeText(int startVal, bool delayed)
        {
            return startVal switch
            {
                0 => "引导启动",
                1 => "系统启动",
                2 => delayed ? "自动(延迟)" : "自动",
                3 => "手动",
                4 => "已禁用",
                _ => "其他"
            };
        }

        private static string ResolveDescriptionString(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "";
            if (!raw.StartsWith("@")) return raw;

            try
            {
                var sb = new StringBuilder(1024);
                if (SHLoadIndirectString(raw, sb, sb.Capacity, IntPtr.Zero) == 0)
                    return sb.ToString();
            }
            catch { }
            return "";
        }

        private static string ExtractExecutablePath(string rawImagePath)
        {
            if (string.IsNullOrWhiteSpace(rawImagePath)) return "";
            string s = rawImagePath.Trim();

            if (s.StartsWith(@"\SystemRoot\", StringComparison.OrdinalIgnoreCase))
                s = @"%SystemRoot%\" + s.Substring(@"\SystemRoot\".Length);

            s = Environment.ExpandEnvironmentVariables(s).Trim();

            string path;
            if (s.StartsWith("\""))
            {
                int endQuote = s.IndexOf('"', 1);
                path = endQuote > 1 ? s.Substring(1, endQuote - 1) : s.Trim('"');
            }
            else
            {
                int exeIdx = s.IndexOf(".exe ", StringComparison.OrdinalIgnoreCase);
                if (exeIdx > 0)
                {
                    path = s.Substring(0, exeIdx + 4);
                }
                else if (s.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    path = s;
                }
                else
                {
                    int sp = s.IndexOf(' ');
                    path = sp > 0 ? s.Substring(0, sp) : s;
                }
            }

            if (!Path.IsPathRooted(path) && !string.IsNullOrEmpty(path))
            {
                string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
                string candidate = Path.Combine(winDir, path);
                if (File.Exists(candidate)) path = candidate;
            }

            return path;
        }

        private static (string label, bool isThirdParty) DetectServiceSource(string exePath)
        {
            if (string.IsNullOrEmpty(exePath))
                return ("未知", false);

            string lower = exePath.ToLowerInvariant();
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows).ToLowerInvariant();
            string sys32 = Path.Combine(winDir, "system32");

            bool inWindowsCore = lower.StartsWith(sys32) || lower.StartsWith(Path.Combine(winDir, "syswow64")) ||
                                 lower.StartsWith(Path.Combine(winDir, "winsxs")) || lower.StartsWith(Path.Combine(winDir, "servicing"));

            string company = "";
            if (File.Exists(exePath))
            {
                try
                {
                    var fvi = FileVersionInfo.GetVersionInfo(exePath);
                    company = fvi.CompanyName?.Trim() ?? "";
                }
                catch { }
            }

            if (inWindowsCore || company.IndexOf("Microsoft", StringComparison.OrdinalIgnoreCase) >= 0)
                return ("微软 / 系统", false);

            if (!string.IsNullOrEmpty(company))
            {
                if (company.Length > 14) company = company.Substring(0, 14) + "…";
                return ($"第三方 ({company})", true);
            }

            return ("第三方软件", true);
        }

        private static int QueryServicePid(IntPtr hScm, string serviceName)
        {
            IntPtr hSvc = OpenService(hScm, serviceName, SERVICE_QUERY_STATUS);
            if (hSvc == IntPtr.Zero) return 0;

            IntPtr buf = Marshal.AllocHGlobal(36);
            try
            {
                if (QueryServiceStatusEx(hSvc, SC_STATUS_PROCESS_INFO, buf, 36, out _))
                {
                    var ssp = Marshal.PtrToStructure<SERVICE_STATUS_PROCESS>(buf);
                    return ssp.dwProcessId;
                }
            }
            catch { }
            finally
            {
                Marshal.FreeHGlobal(buf);
                CloseServiceHandle(hSvc);
            }
            return 0;
        }

        private static double GetProcessMemoryMb(int pid)
        {
            try
            {
                using var proc = Process.GetProcessById(pid);
                return proc.WorkingSet64 / (1024.0 * 1024.0);
            }
            catch
            {
                return 0;
            }
        }
    }
}
