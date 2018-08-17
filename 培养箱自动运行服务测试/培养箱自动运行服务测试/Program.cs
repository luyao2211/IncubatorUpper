using Newtonsoft.Json;
using OPCAutomation;
using SocketSerialTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CameraService;
using ImageSample_1;

namespace 培养箱自动运行服务测试
{
    class Program
    {
        //OPC参数
        static string strHostIP = "";
        static string strHostName = "";
        static OPCServer KepServer;
        static List<string> cmbServerName = new List<string>();
        static bool opc_connected = false;
        static List<string> listBox1 = new List<string>();
        static OPCGroups KepGroups;
        static OPCGroup KepGroup;
        static OPCItems KepItems;
        static int itmHandleClient = 0;
        static int itmHandleServer = 0;
        static OPCItem KepItem;

        //串口参数
        static List<string> sl = new List<string>();
        static private System.IO.Ports.SerialPort serialPort = new System.IO.Ports.SerialPort();
        static string received = "";

        //机器视觉与顶空分析计数
        static int SampleCount = 0;
        //监听变量
        static int? selectedAddress = null;
        //操作指令
        static string Command = "";
        //当前正在进行的操作详细信息
        static OperationInfo operationInfo;
        //内外盘原点位置
        static int OutOrigin = 0;
        static int InOrigin = 0;
        //上下层有无培养器以及培养器信息
        static bool Upper;
        static bool Lower;
        static ResIncubatorInfo UpperresIncubatorInfo;
        static ResIncubatorInfo LowerresIncubatorInfo;
        //机器视觉存储文件夹
        static string PhotoDir = "D:\\机器视觉系统\\样品采集_" + DateTime.Now.ToString("yyyyMMdd") + "\\";

        static void Main(string[] args)
        {
            //服务初始化
            GetLocalServer();
            btnConnLocalServer_Click();
            init_Serial_List();
            serialPort.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(serialPort_DataReceived);
            listBox1_SelectedIndexChanged("通道 1.设备 1.外盘原点位置HMI");
            while (OutOrigin == 0)
            {
                Thread.Sleep(1000);
            }
            listBox1_SelectedIndexChanged("通道 1.设备 1.内盘原点位置HMI");
            while (InOrigin == 0)
            {
                Thread.Sleep(1000);
            }

            while (0 != 1)
            {
                if (Command == "")
                {
                    if (DateTime.Now.Hour == 9 && DateTime.Now.Minute == 55 && DateTime.Now.Second < 50)
                    {
                        PhotoDir = "D:\\机器视觉系统\\样品采集_" + DateTime.Now.ToString("yyyyMMdd") + "\\";
                        Command = "Test";
                        operationInfo = new OperationInfo()
                        {
                            EquipmentId = "Incu_001",
                            OperationTime = DateTime.Now,
                            OperationCode = "OP002",
                            OperationValue = "",
                            OperationResult = "",
                            TerminalIP = strHostIP,
                            TerminalName = strHostName,
                            revUserId = "Shangweiji"
                        };
                        string strURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentSetData";
                        Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(operationInfo)));
                        Test_Incubator();
                    }
                    else
                    {
                        string strURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentGetEquipmentOpsByAnyProperty";
                        string strBody = "{\"EquipmentId\": \"Incu_001\",\"OperationTimeS\": \"" + DateTime.Now.AddSeconds(-5).ToString("s") + "\",\"OperationTimeE\": \"" + DateTime.Now.ToString("s") + "\",\"OperationCode\": \"OP202\",\"OperationValue\": null,\"OperationResult\": null,\"ReDateTimeS\": null,\"ReDateTimeE\": null,\"ReTerminalIP\": null,\"ReTerminalName\": null,\"ReUserId\": null,\"ReIdentify\": null,\"GetOperationCode\": 1,\"GetOperationValue\": 1,\"GetOperationResult\": 1,\"GetRevisionInfo\": 1}";
                        List<string> strTests = HttpPost(strURL, strBody).Split('}').ToList();
                        strTests.Remove("]");
                        if (strTests[0] != "[]")
                        {
                            operationInfo = JsonHelper.DeserializeJsonToObject<OperationInfo>(strTests[0].Substring(1, strTests[0].Length - 1) + "}");
                            Command = "Reset";
                            Reset_Incubator();
                        }
                    }
                }
                Thread.Sleep(1000);
            }
        } 

        /// <summary>
        /// 初始化
        /// </summary>
        static private void Reset_Incubator()
        {
            listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服回的点HMI");
            btnWrite_Click("1");
            Thread.Sleep(1000);
            btnWrite_Click("0");
            Thread.Sleep(1000);
            //模拟，此时“相机伺服原点搜索完成”应为false
            //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服原点搜索完成");
            //btnWrite_Click("0");
            //模拟结束
            listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服原点搜索完成");
            return;
        }

        /// <summary>
        /// 检测
        /// </summary>
        static private void Test_Incubator()
        {
            listBox1_SelectedIndexChanged("通道 1.设备 1.外盘回原点HMI");
            btnWrite_Click("1");
            Thread.Sleep(1000);
            btnWrite_Click("0");
            Thread.Sleep(1000);
            //模拟，此时“外盘原点搜索完成”应为false
            //listBox1_SelectedIndexChanged("通道 1.设备 1.外盘原点搜索完成");
            //btnWrite_Click("0");
            //模拟结束
            listBox1_SelectedIndexChanged("通道 1.设备 1.外盘当前位置");
            return;
        }

        /// <summary>
        /// 获取本地OPC服务器列表
        /// </summary>
        static private void GetLocalServer()
        {
            //获取本地计算机IP,计算机名称
            IPHostEntry IPHost = Dns.Resolve(Environment.MachineName);
            if (IPHost.AddressList.Length > 0)
            {
                strHostIP = IPHost.AddressList[0].ToString();
            }
            else
            {
                return;
            }
            //通过IP来获取计算机名称，可用在局域网内
            IPHostEntry ipHostEntry = Dns.GetHostByAddress(strHostIP);
            strHostName = ipHostEntry.HostName.ToString();
            //获取本地计算机上的OPCServerName
            try
            {
                KepServer = new OPCServer();
                object serverList = KepServer.GetOPCServers(strHostName);

                foreach (string turn in (Array)serverList)
                {
                    cmbServerName.Add(turn);
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("枚举本地OPC服务器出错：" + err.Message);
            }
            return;
        }

        /// <summary>
        /// 连接ＯＰＣ服务器
        /// </summary>
        static private void btnConnLocalServer_Click()
        {
            try
            {
                if (!ConnectRemoteServer("", cmbServerName[0]))
                {
                    return;
                }
                opc_connected = true;
                GetServerInfo();
                RecurBrowse(KepServer.CreateBrowser());
                if (!CreateGroup())
                {
                    return;
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("初始化出错：" + err.Message);
            }
            return;
        }

        /// <summary>
        /// 连接OPC服务器
        /// </summary>
        /// <param name="remoteServerIP">OPCServerIP</param>
        /// <param name="remoteServerName">OPCServer名称</param>
        static private bool ConnectRemoteServer(string remoteServerIP, string remoteServerName)
        {
            try
            {
                KepServer.Connect(remoteServerName, remoteServerIP);
                if (KepServer.ServerState == (int)OPCServerState.OPCRunning)
                {
                    Console.Write("已连接到-" + KepServer.ServerName + "\t");
                }
                else
                {
                    //这里你可以根据返回的状态来自定义显示信息，请查看自动化接口API文档
                    Console.Write("状态：" + KepServer.ServerState.ToString() + "\t");
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("连接远程服务器出现错误：" + err.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 获取服务器信息，并显示在窗体状态栏上
        /// </summary>
        static private void GetServerInfo()
        {
            Console.Write("开始时间:" + KepServer.StartTime.ToString() + "\t");
            Console.WriteLine("版本:" + KepServer.MajorVersion.ToString() + "." + KepServer.MinorVersion.ToString() + "." + KepServer.BuildNumber.ToString());
            return;
        }

        /// <summary>
        /// 列出OPC服务器中所有节点
        /// </summary>
        /// <param name="oPCBrowser"></param>
        static private void RecurBrowse(OPCBrowser oPCBrowser)
        {
            //展开分支
            oPCBrowser.ShowBranches();
            //展开叶子
            oPCBrowser.ShowLeafs(true);
            foreach (object turn in oPCBrowser)
            {
                listBox1.Add(turn.ToString());
            }
            return;
        }

        /// <summary>
        /// 创建组
        /// </summary>
        static private bool CreateGroup()
        {
            try
            {
                KepGroups = KepServer.OPCGroups;
                KepGroup = KepGroups.Add("OPCDOTNETGROUP");
                SetGroupProperty();
                KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(KepGroup_DataChange);
                KepGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(KepGroup_AsyncWriteComplete);
                KepItems = KepGroup.OPCItems;
            }
            catch (Exception err)
            {
                Console.WriteLine("创建组出现错误：" + err.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 设置组属性
        /// </summary>
        static private void SetGroupProperty()
        {
            KepServer.OPCGroups.DefaultGroupIsActive = Convert.ToBoolean("true");
            KepServer.OPCGroups.DefaultGroupDeadband = Convert.ToInt32("0");
            KepGroup.UpdateRate = Convert.ToInt32("250");
            KepGroup.IsActive = Convert.ToBoolean("true");
            KepGroup.IsSubscribed = Convert.ToBoolean("true");
            return;
        }

        /// <summary>
        /// 每当项数据有变化时执行的事件
        /// </summary>
        /// <param name="TransactionID">处理ID</param>
        /// <param name="NumItems">项个数</param>
        /// <param name="ClientHandles">项客户端句柄</param>
        /// <param name="ItemValues">TAG值</param>
        /// <param name="Qualities">品质</param>
        /// <param name="TimeStamps">时间戳</param>
        static void KepGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            //为了测试，所以加了控制台的输出，来查看事物ID号
            //Console.WriteLine("********"+TransactionID.ToString()+"*********");
            if (selectedAddress == 14)
            {
                OutOrigin = Convert.ToInt32(ItemValues.GetValue(1));
            }
            if (selectedAddress == 25)
            {
                InOrigin = Convert.ToInt32(ItemValues.GetValue(1));
            }
            if (ItemValues.GetValue(1).ToString() == "True" || Convert.ToInt32(ItemValues.GetValue(1)) == OutOrigin || Convert.ToInt32(ItemValues.GetValue(1)) == InOrigin)
            {
                if (Command == "Reset")
                {
                    switch (selectedAddress)
                    {
                        case 3:
                            //模拟，此时“相机伺服到外盘位置HMI”应为true
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到外盘位置HMI");
                            //btnWrite_Click("1");
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到内盘位置HMI");
                            //btnWrite_Click("0");
                            //模拟结束
                            Thread.Sleep(1000);
                            listBox1_SelectedIndexChanged("通道 1.设备 1.外盘回原点HMI");
                            btnWrite_Click("1");
                            Thread.Sleep(1000);
                            btnWrite_Click("0");
                            Thread.Sleep(1000);
                            //模拟，此时“外盘原点搜索完成”应为false
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.外盘原点搜索完成");
                            //btnWrite_Click("0");
                            //模拟结束
                            listBox1_SelectedIndexChanged("通道 1.设备 1.外盘当前位置");
                            break;
                        case 20:
                            Thread.Sleep(1000);
                            listBox1_SelectedIndexChanged("通道 1.设备 1.相机内盘位置HMI");
                            btnWrite_Click("1");
                            Thread.Sleep(1000);
                            btnWrite_Click("0");
                            Thread.Sleep(1000);
                            //模拟，此时“相机伺服到外盘位置HMI”应为false
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到外盘位置HMI");
                            //btnWrite_Click("0");
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到内盘位置HMI");
                            //btnWrite_Click("0");
                            //模拟结束
                            listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到内盘位置HMI");
                            break;
                        case 6:
                            Thread.Sleep(1000);
                            listBox1_SelectedIndexChanged("通道 1.设备 1.内盘回原点HMI");
                            btnWrite_Click("1");
                            Thread.Sleep(1000);
                            btnWrite_Click("0");
                            Thread.Sleep(1000);
                            //模拟，此时“内盘原点搜索完成”应为false
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.内盘原点搜索完成");
                            //btnWrite_Click("0");
                            //模拟结束
                            listBox1_SelectedIndexChanged("通道 1.设备 1.内盘当前位置");
                            break;
                        case 31:
                            Thread.Sleep(1000);
                            listBox1_SelectedIndexChanged("通道 1.设备 1.相机外盘位置HMI");
                            btnWrite_Click("1");
                            Thread.Sleep(1000);
                            btnWrite_Click("0");
                            Thread.Sleep(1000);
                            //模拟，此时“相机伺服到内盘位置HMI”应为false
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到外盘位置HMI");
                            //btnWrite_Click("0");
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到内盘位置HMI");
                            //btnWrite_Click("0");
                            //模拟结束
                            listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到外盘位置HMI");
                            break;
                        case 5:
                            Command = "";
                            operationInfo.OperationResult = "ok";
                            string strURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentSetData";
                            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(operationInfo)));
                            break;
                    }
                }
                if (Command == "Test")
                {
                    switch (selectedAddress)
                    {
                        case 20:
                            Thread.Sleep(1000);
                            listBox1_SelectedIndexChanged("通道 1.设备 1.外盘单动HMI");
                            btnWrite_Click("1");
                            Thread.Sleep(1000);
                            btnWrite_Click("0");
                            Thread.Sleep(1000);
                            //模拟，此时“外盘定位完成”应为false
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.外盘定位完成");
                            //btnWrite_Click("0");
                            //模拟结束
                            listBox1_SelectedIndexChanged("通道 1.设备 1.外盘定位完成");
                            break;
                        case 19:
                            SampleCount += 1;
                            if (SampleCount == 30)
                            {
                                SampleCount = 0;
                                listBox1_SelectedIndexChanged("通道 1.设备 1.相机内盘位置HMI");
                                btnWrite_Click("1");
                                Thread.Sleep(1000);
                                btnWrite_Click("0");
                                Thread.Sleep(1000);
                                //模拟，此时“相机伺服到外盘位置HMI”应为false
                                //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到外盘位置HMI");
                                //btnWrite_Click("0");
                                //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到内盘位置HMI");
                                //btnWrite_Click("0");
                                //模拟结束
                                listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到内盘位置HMI");
                            }
                            else
                            {
                                string strURL = "http://121.43.107.106:8063/Api/v1/Result/ResIncubatorGetResultTubesByAnyProperty";
                                string strBody = "{\"TestId\": null,\"TubeNo\": null,\"CultureId\": null,\"BacterId\": null,\"OtherRea\": null,\"IncubatorId\": \"Incu_001\",\"Place\": \"Upper_Out_" + SampleCount.ToString() + "\",\"StartTimeS\": null,\"StartTimeE\": null,\"EndTimeS\": null,\"EndTimeE\": null,\"AnalResult\": null,\"PutinPeople\": null,\"PutoutPeople\": null,\"PutoutTimeS\": null,\"PutoutTimeE\": null,\"GetCultureId\": 1,\"GetBacterId\": 1,\"GetOtherRea\": 1,\"GetIncubatorId\": 1,\"GetPlace\": 1,\"GetStartTime\": 1,\"GetEndTime\": 1,\"GetAnalResult\": 1,\"GetPutinPeople\": 1,\"GetPutoutPeople\": 1,\"GetPutoutTime\": 1}";
                                List<string> strTests = HttpPost(strURL, strBody).Split('}').ToList();
                                strTests.Remove("]");
                                if (strTests[0] != "[]")
                                {
                                    Upper = true;
                                    UpperresIncubatorInfo = JsonHelper.DeserializeJsonToObject<ResIncubatorInfo>(strTests[0].Substring(1, strTests[0].Length - 1) + "}");
                                }
                                else
                                {
                                    Upper = false;
                                }
                                strBody = "{\"TestId\": null,\"TubeNo\": null,\"CultureId\": null,\"BacterId\": null,\"OtherRea\": null,\"IncubatorId\": \"Incu_001\",\"Place\": \"Lower_Out_" + SampleCount.ToString() + "\",\"StartTimeS\": null,\"StartTimeE\": null,\"EndTimeS\": null,\"EndTimeE\": null,\"AnalResult\": null,\"PutinPeople\": null,\"PutoutPeople\": null,\"PutoutTimeS\": null,\"PutoutTimeE\": null,\"GetCultureId\": 1,\"GetBacterId\": 1,\"GetOtherRea\": 1,\"GetIncubatorId\": 1,\"GetPlace\": 1,\"GetStartTime\": 1,\"GetEndTime\": 1,\"GetAnalResult\": 1,\"GetPutinPeople\": 1,\"GetPutoutPeople\": 1,\"GetPutoutTime\": 1}";
                                strTests = HttpPost(strURL, strBody).Split('}').ToList();
                                strTests.Remove("]");
                                if (strTests[0] != "[]")
                                {
                                    Lower = true;
                                    LowerresIncubatorInfo = JsonHelper.DeserializeJsonToObject<ResIncubatorInfo>(strTests[0].Substring(1, strTests[0].Length - 1) + "}");
                                }
                                else
                                {
                                    Lower = false;
                                }
                                if (Upper == true)
                                {
                                    //相机机器视觉，李润泽
                                    CameraServiceAPI cam = new CameraServiceAPI();
                                    string PhotoAddress = PhotoDir + UpperresIncubatorInfo.TestId + UpperresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg";
                                    cam.Shoot_One_Slot(PhotoAddress, "", true, PhotoDir);
                                    //ftp上传
                                    FtpHelper ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                    FileInfo image = new FileInfo(@PhotoAddress);
                                    bool shangchuan = ftpHelper.Upload(image, "\\SterilityWebAPI\\Image\\" + UpperresIncubatorInfo.TestId + UpperresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg");
                                    PhotoPath photoPath = new PhotoPath()
                                    {
                                        path = PhotoAddress
                                    };
                                    strURL = "http://localhost:5000/predictbyfilename";
                                    //插入机器视觉
                                    string[] VisualList = HttpPost(strURL, JsonHelper.SerializeObject(photoPath)).Split('"');
                                    if (VisualList.Length == 11)
                                    {
                                        string outcome = "";
                                        if (VisualList.GetValue(3).ToString() == "All_of_Negative")
                                        {
                                            outcome = "无菌";
                                        }
                                        else
                                        {
                                            outcome = "有菌：" + VisualList.GetValue(3);
                                        }
                                        TestPicture testPicture = new TestPicture()
                                        {
                                            TestId = UpperresIncubatorInfo.TestId,
                                            TubeNo = UpperresIncubatorInfo.TubeNo,
                                            PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                            CameraTime = DateTime.Now,
                                            ImageAddress = "/Image/" + UpperresIncubatorInfo.TestId + UpperresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg",
                                            AnalResult = outcome
                                        };
                                        strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTestPictureSetData";
                                        Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(testPicture)));
                                    }
                                    //顶空分析
                                    open_Serial(2);
                                    while (1 != 0)
                                    {
                                        Thread.Sleep(1000);
                                        if (received != "")
                                        {
                                            //插入顶空分析
                                            TopAnalysis topAnalysis = new TopAnalysis()
                                            {
                                                TestId = UpperresIncubatorInfo.TestId,
                                                TubeNo = UpperresIncubatorInfo.TubeNo,
                                                PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                                CameraTime = DateTime.Now,
                                                AnalResult = received
                                            };
                                            strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTopAnalysisSetData";
                                            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(topAnalysis)));
                                            received = "";
                                            close_Serial();
                                            break;
                                        }
                                    }
                                }
                                if (Lower == true)
                                {
                                    //相机机器视觉，李润泽
                                    CameraServiceAPI cam = new CameraServiceAPI();
                                    string PhotoAddress = PhotoDir + LowerresIncubatorInfo.TestId + LowerresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg";
                                    cam.Shoot_One_Slot(PhotoAddress, "", false, PhotoDir);
                                    //ftp上传
                                    FtpHelper ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                    FileInfo image = new FileInfo(@PhotoAddress);
                                    bool shangchuan = ftpHelper.Upload(image, "\\SterilityWebAPI\\Image\\" + LowerresIncubatorInfo.TestId + LowerresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg");
                                    PhotoPath photoPath = new PhotoPath()
                                    {
                                        path = PhotoAddress
                                    };
                                    strURL = "http://localhost:5000/predictbyfilename";
                                    //插入机器视觉
                                    string[] VisualList = HttpPost(strURL, JsonHelper.SerializeObject(photoPath)).Split('"');
                                    if (VisualList.Length == 11)
                                    {
                                        string outcome = "";
                                        if (VisualList.GetValue(3).ToString() == "All_of_Negative")
                                        {
                                            outcome = "无菌";
                                        }
                                        else
                                        {
                                            outcome = "有菌：" + VisualList.GetValue(3);
                                        }
                                        TestPicture testPicture = new TestPicture()
                                        {
                                            TestId = LowerresIncubatorInfo.TestId,
                                            TubeNo = LowerresIncubatorInfo.TubeNo,
                                            PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                            CameraTime = DateTime.Now,
                                            ImageAddress = "/Image/" + LowerresIncubatorInfo.TestId + LowerresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg",
                                            AnalResult = outcome
                                        };
                                        strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTestPictureSetData";
                                        Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(testPicture)));
                                    }
                                    //顶空分析
                                    open_Serial(1);
                                    while (1 != 0)
                                    {
                                        Thread.Sleep(1000);
                                        if (received != "")
                                        {
                                            //插入顶空分析
                                            TopAnalysis topAnalysis = new TopAnalysis()
                                            {
                                                TestId = LowerresIncubatorInfo.TestId,
                                                TubeNo = LowerresIncubatorInfo.TubeNo,
                                                PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                                CameraTime = DateTime.Now,
                                                AnalResult = received
                                            };
                                            strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTopAnalysisSetData";
                                            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(topAnalysis)));
                                            received = "";
                                            close_Serial();
                                            break;
                                        }
                                    }
                                }
                                listBox1_SelectedIndexChanged("通道 1.设备 1.外盘单动HMI");
                                btnWrite_Click("1");
                                Thread.Sleep(1000);
                                btnWrite_Click("0");
                                Thread.Sleep(1000);
                                //模拟，此时“外盘定位完成”应为false
                                //listBox1_SelectedIndexChanged("通道 1.设备 1.外盘定位完成");
                                //btnWrite_Click("0");
                                //模拟结束
                                listBox1_SelectedIndexChanged("通道 1.设备 1.外盘定位完成");
                            }
                            break;
                        case 6:
                            Thread.Sleep(1000);
                            listBox1_SelectedIndexChanged("通道 1.设备 1.内盘回原点HMI");
                            btnWrite_Click("1");
                            Thread.Sleep(1000);
                            btnWrite_Click("0");
                            Thread.Sleep(1000);
                            //模拟，此时“内盘原点搜索完成”应为false
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.内盘原点搜索完成");
                            //btnWrite_Click("0");
                            //模拟结束
                            listBox1_SelectedIndexChanged("通道 1.设备 1.内盘当前位置");
                            break;
                        case 31:
                            Thread.Sleep(1000);
                            listBox1_SelectedIndexChanged("通道 1.设备 1.内盘单动HMI");
                            btnWrite_Click("1");
                            Thread.Sleep(1000);
                            btnWrite_Click("0");
                            Thread.Sleep(1000);
                            //模拟，此时“内盘定位完成”应为false
                            //listBox1_SelectedIndexChanged("通道 1.设备 1.内盘定位完成");
                            //btnWrite_Click("0");
                            //模拟结束
                            listBox1_SelectedIndexChanged("通道 1.设备 1.内盘定位完成");
                            break;
                        case 30:
                            SampleCount += 1;
                            if (SampleCount == 15)
                            {
                                SampleCount = 0;
                                listBox1_SelectedIndexChanged("通道 1.设备 1.相机外盘位置HMI");
                                btnWrite_Click("1");
                                Thread.Sleep(1000);
                                btnWrite_Click("0");
                                Thread.Sleep(1000);
                                //模拟，此时“相机伺服到内盘位置HMI”应为false
                                //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到外盘位置HMI");
                                //btnWrite_Click("0");
                                //listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到内盘位置HMI");
                                //btnWrite_Click("0");
                                //模拟结束
                                listBox1_SelectedIndexChanged("通道 1.设备 1.相机伺服到外盘位置HMI");
                            }
                            else
                            {
                                string strURL = "http://121.43.107.106:8063/Api/v1/Result/ResIncubatorGetResultTubesByAnyProperty";
                                string strBody = "{\"TestId\": null,\"TubeNo\": null,\"CultureId\": null,\"BacterId\": null,\"OtherRea\": null,\"IncubatorId\": \"Incu_001\",\"Place\": \"Upper_In_" + SampleCount.ToString() + "\",\"StartTimeS\": null,\"StartTimeE\": null,\"EndTimeS\": null,\"EndTimeE\": null,\"AnalResult\": null,\"PutinPeople\": null,\"PutoutPeople\": null,\"PutoutTimeS\": null,\"PutoutTimeE\": null,\"GetCultureId\": 1,\"GetBacterId\": 1,\"GetOtherRea\": 1,\"GetIncubatorId\": 1,\"GetPlace\": 1,\"GetStartTime\": 1,\"GetEndTime\": 1,\"GetAnalResult\": 1,\"GetPutinPeople\": 1,\"GetPutoutPeople\": 1,\"GetPutoutTime\": 1}";
                                List<string> strTests = HttpPost(strURL, strBody).Split('}').ToList();
                                strTests.Remove("]");
                                if (strTests[0] != "[]")
                                {
                                    Upper = true;
                                    UpperresIncubatorInfo = JsonHelper.DeserializeJsonToObject<ResIncubatorInfo>(strTests[0].Substring(1, strTests[0].Length - 1) + "}");
                                }
                                else
                                {
                                    Upper = false;
                                }
                                strBody = "{\"TestId\": null,\"TubeNo\": null,\"CultureId\": null,\"BacterId\": null,\"OtherRea\": null,\"IncubatorId\": \"Incu_001\",\"Place\": \"Lower_In_" + SampleCount.ToString() + "\",\"StartTimeS\": null,\"StartTimeE\": null,\"EndTimeS\": null,\"EndTimeE\": null,\"AnalResult\": null,\"PutinPeople\": null,\"PutoutPeople\": null,\"PutoutTimeS\": null,\"PutoutTimeE\": null,\"GetCultureId\": 1,\"GetBacterId\": 1,\"GetOtherRea\": 1,\"GetIncubatorId\": 1,\"GetPlace\": 1,\"GetStartTime\": 1,\"GetEndTime\": 1,\"GetAnalResult\": 1,\"GetPutinPeople\": 1,\"GetPutoutPeople\": 1,\"GetPutoutTime\": 1}";
                                strTests = HttpPost(strURL, strBody).Split('}').ToList();
                                strTests.Remove("]");
                                if (strTests[0] != "[]")
                                {
                                    Lower = true;
                                    LowerresIncubatorInfo = JsonHelper.DeserializeJsonToObject<ResIncubatorInfo>(strTests[0].Substring(1, strTests[0].Length - 1) + "}");
                                }
                                else
                                {
                                    Lower = false;
                                }
                                if (Upper == true)
                                {
                                    //相机机器视觉，李润泽
                                    CameraServiceAPI cam = new CameraServiceAPI();
                                    string PhotoAddress = PhotoDir + UpperresIncubatorInfo.TestId + UpperresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg";
                                    cam.Shoot_One_Slot(PhotoAddress, "", true, PhotoDir);
                                    //ftp上传
                                    FtpHelper ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                    FileInfo image = new FileInfo(@PhotoAddress);
                                    bool shangchuan = ftpHelper.Upload(image, "\\SterilityWebAPI\\Image\\" + UpperresIncubatorInfo.TestId + UpperresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg");
                                    PhotoPath photoPath = new PhotoPath()
                                    {
                                        path = PhotoAddress
                                    };
                                    strURL = "http://localhost:5000/predictbyfilename";
                                    //插入机器视觉
                                    string[] VisualList = HttpPost(strURL, JsonHelper.SerializeObject(photoPath)).Split('"');
                                    if (VisualList.Length == 11)
                                    {
                                        string outcome = "";
                                        if (VisualList.GetValue(3).ToString() == "All_of_Negative")
                                        {
                                            outcome = "无菌";
                                        }
                                        else
                                        {
                                            outcome = "有菌：" + VisualList.GetValue(3);
                                        }
                                        TestPicture testPicture = new TestPicture()
                                        {
                                            TestId = UpperresIncubatorInfo.TestId,
                                            TubeNo = UpperresIncubatorInfo.TubeNo,
                                            PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                            CameraTime = DateTime.Now,
                                            ImageAddress = "/Image/" + UpperresIncubatorInfo.TestId + UpperresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg",
                                            AnalResult = outcome
                                        };
                                        strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTestPictureSetData";
                                        Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(testPicture)));
                                    }
                                    //顶空分析
                                    open_Serial(2);
                                    while (1 != 0)
                                    {
                                        Thread.Sleep(1000);
                                        if (received != "")
                                        {
                                            //插入顶空分析
                                            TopAnalysis topAnalysis = new TopAnalysis()
                                            {
                                                TestId = UpperresIncubatorInfo.TestId,
                                                TubeNo = UpperresIncubatorInfo.TubeNo,
                                                PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                                CameraTime = DateTime.Now,
                                                AnalResult = received
                                            };
                                            strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTopAnalysisSetData";
                                            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(topAnalysis)));
                                            received = "";
                                            close_Serial();
                                            break;
                                        }
                                    }
                                }
                                if (Lower == true)
                                {
                                    //相机机器视觉，李润泽
                                    CameraServiceAPI cam = new CameraServiceAPI();
                                    string PhotoAddress = PhotoDir + LowerresIncubatorInfo.TestId + LowerresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg";
                                    cam.Shoot_One_Slot(PhotoAddress, "", false, PhotoDir);
                                    //ftp上传
                                    FtpHelper ftpHelper = new FtpHelper("121.43.107.106", "administrator", "qaz@163.com");
                                    FileInfo image = new FileInfo(@PhotoAddress);
                                    bool shangchuan = ftpHelper.Upload(image, "\\SterilityWebAPI\\Image\\" + LowerresIncubatorInfo.TestId + LowerresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg");
                                    PhotoPath photoPath = new PhotoPath()
                                    {
                                        path = PhotoAddress
                                    };
                                    strURL = "http://localhost:5000/predictbyfilename";
                                    Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(photoPath)));
                                    //插入机器视觉
                                    string[] VisualList = HttpPost(strURL, JsonHelper.SerializeObject(photoPath)).Split('"');
                                    if (VisualList.Length == 11)
                                    {
                                        string outcome = "";
                                        if (VisualList.GetValue(3).ToString() == "All_of_Negative")
                                        {
                                            outcome = "无菌";
                                        }
                                        else
                                        {
                                            outcome = "有菌：" + VisualList.GetValue(3);
                                        }
                                        TestPicture testPicture = new TestPicture()
                                        {
                                            TestId = LowerresIncubatorInfo.TestId,
                                            TubeNo = LowerresIncubatorInfo.TubeNo,
                                            PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                            CameraTime = DateTime.Now,
                                            ImageAddress = "/Image/" + LowerresIncubatorInfo.TestId + LowerresIncubatorInfo.TubeNo + DateTime.Now.ToString("yyyy-MM-dd-HH") + ".jpg",
                                            AnalResult = outcome
                                        };
                                        strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTestPictureSetData";
                                        Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(testPicture)));
                                    }
                                    //顶空分析
                                    open_Serial(1);
                                    while (1 != 0)
                                    {
                                        Thread.Sleep(1000);
                                        if (received != "")
                                        {
                                            //插入顶空分析
                                            TopAnalysis topAnalysis = new TopAnalysis()
                                            {
                                                TestId = LowerresIncubatorInfo.TestId,
                                                TubeNo = LowerresIncubatorInfo.TubeNo,
                                                PictureId = DateTime.Now.ToString("yyyyMMddHH"),
                                                CameraTime = DateTime.Now,
                                                AnalResult = received
                                            };
                                            strURL = "http://121.43.107.106:8063/Api/v1/Result/ResTopAnalysisSetData";
                                            Console.WriteLine(HttpPost(strURL, JsonHelper.SerializeObject(topAnalysis)));
                                            received = "";
                                            close_Serial();
                                            break;
                                        }
                                    }
                                }
                                listBox1_SelectedIndexChanged("通道 1.设备 1.内盘单动HMI");
                                btnWrite_Click("1");
                                Thread.Sleep(1000);
                                btnWrite_Click("0");
                                Thread.Sleep(1000);
                                //模拟，此时“内盘定位完成”应为false
                                //listBox1_SelectedIndexChanged("通道 1.设备 1.内盘定位完成");
                                //btnWrite_Click("0");
                                //模拟结束
                                listBox1_SelectedIndexChanged("通道 1.设备 1.内盘定位完成");
                            }
                            break;
                        case 5:
                            Command = "";
                            operationInfo.OperationResult = "ok";
                            string StrURL = "http://121.43.107.106:8063/Api/v1/Operation/OpEquipmentSetData";
                            Console.WriteLine(HttpPost(StrURL, JsonHelper.SerializeObject(operationInfo)));
                            break;
                    }
                }
            }
            return;
        }

        /// <summary>
        /// 写入TAG值时执行的事件
        /// </summary>
        /// <param name="TransactionID"></param>
        /// <param name="NumItems"></param>
        /// <param name="ClientHandles"></param>
        /// <param name="Errors"></param>
        static void KepGroup_AsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
        {
            for (int i = 1; i <= NumItems; i++)
            {
                Console.WriteLine("Tran:" + TransactionID.ToString() + "   CH:" + ClientHandles.GetValue(i).ToString() + "   Error:" + Errors.GetValue(i).ToString());
            }
            return;
        }

        /// <summary>
        /// 选择列表项时处理的事情
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static private void listBox1_SelectedIndexChanged(string tag)
        {
            try
            {
                if (itmHandleClient != 0)
                {
                    Array Errors;
                    OPCItem bItem = KepItems.GetOPCItem(itmHandleServer);
                    //注：OPC中以1为数组的基数
                    int[] temp = new int[2] { 0, bItem.ServerHandle };
                    Array serverHandle = (Array)temp;
                    //移除上一次选择的项
                    KepItems.Remove(KepItems.Count, ref serverHandle, out Errors);
                }
                itmHandleClient = 1234;
                KepItem = KepItems.AddItem(tag, itmHandleClient);
                itmHandleServer = KepItem.ServerHandle;
                selectedAddress = listBox1.FindIndex(s => s == tag);
            }
            catch (Exception err)
            {
                //没有任何权限的项，都是OPC服务器保留的系统项，此处可不做处理。
                itmHandleClient = 0;
                Console.WriteLine("此项为系统保留项:" + err.Message, "提示信息");
            }
            return;
        }

        /// <summary>
        /// 【按钮】写入
        /// </summary>
        static private void btnWrite_Click(string outcome)
        {
            OPCItem bItem = KepItems.GetOPCItem(itmHandleServer);
            int[] temp = new int[2] { 0, bItem.ServerHandle };
            Array serverHandles = (Array)temp;
            object[] valueTemp = new object[2] { "", outcome };
            Array values = (Array)valueTemp;
            Array Errors;
            int cancelID;
            KepGroup.AsyncWrite(1, ref serverHandles, ref values, out Errors, 2009, out cancelID);
            //KepItem.Write(txtWriteTagValue.Text);//这句也可以写入，但并不触发写入事件
            GC.Collect();
            return;
        }

        /// <summary>
        /// 获取串口列表
        /// </summary>
        static private void init_Serial_List()
        {
            sl = SerialPortTool.GetSerialPortList();
            if (sl == null)
            {
                Console.WriteLine("读取串口列表失败");
                return;
            }
            return;
        }

        /// <summary>
        /// 读取串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static private void serialPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            byte[] buffer = new byte[1024];
            int n = serialPort.Read(buffer, 0, 1024);
            received = System.Text.Encoding.UTF8.GetString(buffer, 0, n);
            return;
        }

        /// <summary>
        /// 关闭串口
        /// </summary>
        static private void close_Serial()
        {
            if (serialPort.IsOpen)
            {
                serialPort.Close();
                Console.WriteLine("串口关闭成功");
            }
        }

        /// <summary>
        /// 打开串口
        /// </summary>
        /// <param name="duankouhao"></param>
        /// <returns></returns>
        static private bool open_Serial(int duankouhao)
        {
            if (serialPort.IsOpen)
            {
                return true;
            }
            int baud;
            if (!int.TryParse("9600", out baud))
            {
                return false;
            }
            serialPort.PortName = SerialPortTool.GetSerialPortByName(sl[duankouhao]);
            serialPort.BaudRate = baud;
            try
            {
                serialPort.Open();
            }
            catch (System.IO.IOException ioe)
            {
                Console.WriteLine(ioe.Message);
            }
            catch (System.UnauthorizedAccessException ioe)
            {
                Console.WriteLine(ioe.Message);
                return false;
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
            if (!serialPort.IsOpen)
            {
                Console.WriteLine(serialPort.PortName + ": 打开串口失败");
                return false;
            }
            Console.WriteLine(serialPort.PortName + ": 打开成功, 速率: " + baud);
            return true;
        }

        public class OperationInfo
        {
            public string EquipmentId { get; set; }
            public DateTime OperationTime { get; set; }
            public string OperationCode { get; set; }
            public string OperationValue { get; set; }
            public string OperationResult { get; set; }
            public string TerminalIP { get; set; }
            public string TerminalName { get; set; }
            public string revUserId { get; set; }
        }

        public class ResIncubatorInfo
        {
            public string TestId { get; set; }
            public string TubeNo { get; set; }
            public string CultureId { get; set; }
            public string BacterId { get; set; }
            public string OtherRea { get; set; }
            public string IncubatorId { get; set; }
            public string Place { get; set; }
            public string StartTime { get; set; }
            public string EndTime { get; set; }
            public string AnalResult { get; set; }
            public string PutinPeople { get; set; }
            public string PutoutPeople { get; set; }
            public DateTime PutoutTime { get; set; }
        }

        public class TestPicture
        {
            public string TestId { get; set; }
            public string TubeNo { get; set; }
            public string PictureId { get; set; }
            public DateTime CameraTime { get; set; }
            public string ImageAddress { get; set; }
            public string AnalResult { get; set; }
        }

        public class TopAnalysis
        {
            public string TestId { get; set; }
            public string TubeNo { get; set; }
            public string PictureId { get; set; }
            public DateTime CameraTime { get; set; }
            public string AnalResult { get; set; }
        }

        public class PhotoPath
        {
            public string path { get; set; }
        }

        public static string HttpPost(string strURL, string strBody)
        {
            Encoding encoding = Encoding.UTF8;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Headers.Add("Authorization:Shangweiji");
            byte[] buffer = encoding.GetBytes(strBody);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public class JsonHelper
        {
            /// <summary>
            /// 将对象序列化为JSON格式
            /// </summary>
            /// <param name="o">对象</param>
            /// <returns>json字符串</returns>
            public static string SerializeObject(object o)
            {
                string json = JsonConvert.SerializeObject(o);
                return json;
            }

            /// <summary>
            /// 解析JSON字符串生成对象实体
            /// </summary>
            /// <typeparam name="T">对象类型</typeparam>
            /// <param name="json">json字符串(eg.{"ID":"112","Name":"石子儿"})</param>
            /// <returns>对象实体</returns>
            public static T DeserializeJsonToObject<T>(string json) where T : class
            {
                JsonSerializer serializer = new JsonSerializer();
                StringReader sr = new StringReader(json);
                object o = serializer.Deserialize(new JsonTextReader(sr), typeof(T));
                T t = o as T;
                return t;
            }

            /// <summary>
            /// 解析JSON数组生成对象实体集合
            /// </summary>
            /// <typeparam name="T">对象类型</typeparam>
            /// <param name="json">json数组字符串(eg.[{"ID":"112","Name":"石子儿"}])</param>
            /// <returns>对象实体集合</returns>
            public static List<T> DeserializeJsonToList<T>(string json) where T : class
            {
                JsonSerializer serializer = new JsonSerializer();
                StringReader sr = new StringReader(json);
                object o = serializer.Deserialize(new JsonTextReader(sr), typeof(List<T>));
                List<T> list = o as List<T>;
                return list;
            }

            /// <summary>
            /// 反序列化JSON到给定的匿名对象.
            /// </summary>
            /// <typeparam name="T">匿名对象类型</typeparam>
            /// <param name="json">json字符串</param>
            /// <param name="anonymousTypeObject">匿名对象</param>
            /// <returns>匿名对象</returns>
            public static T DeserializeAnonymousType<T>(string json, T anonymousTypeObject)
            {
                T t = JsonConvert.DeserializeAnonymousType(json, anonymousTypeObject);
                return t;
            }
        }
    }
}
