using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Text;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using static System.Windows.Forms.ListViewItem;

namespace InquireCardTool
{
    public partial class MainForm : Form
    {
        private System.Timers.Timer TimeOutTimer;   //自定義Timer，用於檢查資料接收是否超時
        private Queue<byte> ReceivedQueue = new Queue<byte>();  //接收資料佇列
        private Thread ReceivedThread;  //接收資料執行緒
        private Thread ReaderThread;    //讀卡機流程執行緒
        private MyDelegate Delegate;    //自定義委派事件Class
        iCASH iCASH = new iCASH();
        YHDP YHDP = new YHDP();
        ipass ipass = new ipass();
        EasyCard EasyCard = new EasyCard();
        /*------------------------變數定義-----------------------
         * 1.ReceivedArray:儲存接收資料用陣列
         * 2.CardNumber:儲存卡號用陣列
         * 3.ThreadFlag:執行緒旗標        True:繼續執行     False:停止
         * 4.ReceivedFlag:資料接收旗標    True:資料接收完整 False:未接收到資料或資料不完整
         * 5.TimeOutFlag:資料接收超時旗標 True:超時         False:未超時
         *-------------------------------------------------------*/
        private byte[] ReceivedArray;
        private byte[] CardNumber = new byte[7];
        private bool ThreadFlag = false, ReceivedFlag = false, TimeOutFlag = false;

        public MainForm()
        {
            //畫面初始化
            InitializeComponent();
            //TimeOutTimer初始化設定
            TimeOutTimer = new System.Timers.Timer(3000);
            TimeOutTimer.Elapsed += OnTimedEvent;
            TimeOutTimer.AutoReset = true;
            //UI元件初始化設定
            ComPort_ComboBox.SelectedIndex = 0;
            ComPort_ComboBox.Items.AddRange(SerialPort.GetPortNames());
            //定義委派Class
            Delegate = new MyDelegate(this);
            Console.WriteLine("--------------------------------------------------------------------");
            Console.WriteLine("初始化完成");
        }
        private void ComPort_ComboBox_DropDown(object sender, EventArgs e)
        {
            ComPort_ComboBox.Items.Clear();
            ComPort_ComboBox.Items.Add("選擇ComPort");
            ComPort_ComboBox.Items.AddRange(SerialPort.GetPortNames());
        }
        private void ComPort_ComboBox_TextChanged(object sender, EventArgs e)
        {
            if (RS232.IsOpen)
            {
                CloseReader();
            }
            if (ComPort_ComboBox.Text != "" && ComPort_ComboBox.SelectedIndex != 0)
            {
                RS232.PortName = ComPort_ComboBox.Text;
                OpenReader();
            }
            else
            {
                ComPort_ComboBox.SelectedIndex = 0;
            }

        }
        private void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            TimeOutFlag = true;
            Delegate.UpdateUI("TimeOut，讀卡機無回應" + Environment.NewLine, Info_RichTextBox);
            Console.WriteLine("TimeOut，讀卡機無回應");
        }
        private void Info_RichTextBox_TextChanged(object sender, EventArgs e)
        {
            Info_RichTextBox.SelectionStart = Info_RichTextBox.TextLength;
            Info_RichTextBox.ScrollToCaret();
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ReceivedThread != null && ReaderThread != null)
            {
                if(ReceivedThread.IsAlive || ReaderThread.IsAlive)
                {
                    ThreadFlag = false;
                    while (ReceivedThread.IsAlive != true && ReaderThread.IsAlive != true) ;
                }
            }
        }
        private void OpenReader()
        {
            try
            {
                Delegate.UpdateUI(ComPort_ComboBox, false);
                Delegate.UpdateUI(Content_ListView);
                RS232.Open();
                //Delegate.UpdateUI("串列埠已開啟" + Environment.NewLine, Info_RichTextBox);
                Console.WriteLine("串列埠已開啟");
                ThreadFlag = true;
                ReceivedThread = new Thread(new ThreadStart(DataReceived));
                ReaderThread = new Thread(new ThreadStart(ReaderFunction));
                ReceivedThread.IsBackground = true;
                ReaderThread.IsBackground = true;
                ReceivedThread.Start();
                ReaderThread.Start();
                //Delegate.UpdateUI("執行緒已開啟" + Environment.NewLine, Info_RichTextBox);
                Console.WriteLine("執行緒已開啟");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private void CloseReader()
        {
            try
            {
                Delegate.UpdateUI(ComPort_ComboBox, false);
                ThreadFlag = false;
                Thread.Sleep(1000);
                //Delegate.UpdateUI("執行緒已關閉" + Environment.NewLine, Info_RichTextBox);
                Console.WriteLine("執行緒已關閉");
                RS232.Close();
                //Delegate.UpdateUI("串列埠已關閉" + Environment.NewLine, Info_RichTextBox);
                Console.WriteLine("串列埠已關閉");
                TimeOutTimer.Stop();
                ReceivedFlag = false;
                TimeOutFlag = false;
                Array.Clear(CardNumber, 0, CardNumber.Length);
                if (ReceivedArray != null)
                {
                    Array.Clear(ReceivedArray, 0, ReceivedArray.Length);
                }
                Delegate.UpdateUI(ComPort_ComboBox);
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Delegate.UpdateUI(ComPort_ComboBox, true);
        }
        private void CMDWrite(int CommandCode)
        {
            byte[] Header = { 0xEA };
            byte[] Command = new byte[2];
            byte[] Length;
            byte[] body = { };
            byte[] Tail = { 0x90, 0x00 };

            switch (CommandCode)
            {
                case 0://取得版本
                    {
                        byte[] GetVersionCMD = { 0x01, 0x00 };
                        Command = GetVersionCMD;
                        Array.Resize(ref body, 1);
                        body[0] = 0x00;
                        break;
                    }
                case 1://PPRReset
                    {
                        byte[] PPRResetCMD = { 0x04, 0x01 };
                        Command = PPRResetCMD;
                        Array.Resize(ref body, 71);
                        byte[] CLAToLc = { 0x80, 0x01, 0x00, 0x01,0x40};//5 0
                        byte[] TMLocationID = { 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30, 0x30 };//10 5
                        byte[] TMID = { 0x30, 0x30 };//2 15
                        byte[] TMTXNDateTime= Encoding.ASCII.GetBytes(DateTimeOffset.Now.ToString("yyyyMMddHHmmss"));//14 17
                        byte[] TMSerialNumber = { 0x30, 0x30, 0x30, 0x30, 0x30, 0x31 };//6 31
                        byte[] TMAgentNumber = { 0x30, 0x30, 0x30, 0x30 };//4 37
                        byte[] TXNDateTime = GetUnixTime();//4 41
                        byte[] LocationID = {0x00 };//1 45
                        byte[] NewLocationID = { 0x00, 0x00 };//2 46
                        byte[] ServiceProviderID = { 0x00 };//1 48
                        byte[] NewServiceProviderID = { 0x00, 0x00, 0x00 };//3 49
                        byte[] MicroPaymentFlag = {0x80 };//1 52
                        byte[] OneDayQuotaForMicroPayment = { 0x00, 0x00 };//2 53
                        byte[] OnceQuotaForMicroPayment = { 0x00, 0x00 };//2 55
                        byte[] SAMSlotControlFlag = { 0x11};//1 57
                        byte[] MifareKetSet = { 0x00 };//1 58
                        Array.Copy(CLAToLc,0, body,0, CLAToLc.Length);
                        Array.Copy(TMLocationID, 0, body, 5, TMLocationID.Length);
                        Array.Copy(TMID, 0, body, 15, TMID.Length);
                        Array.Copy(TMTXNDateTime, 0, body, 17, TMTXNDateTime.Length);
                        Array.Copy(TMSerialNumber, 0, body, 31, TMSerialNumber.Length);
                        Array.Copy(TMAgentNumber, 0, body, 37, TMAgentNumber.Length);
                        Array.Copy(TXNDateTime, 0, body, 41, TXNDateTime.Length);
                        Array.Copy(LocationID, 0, body, 45, LocationID.Length);
                        Array.Copy(NewLocationID, 0, body, 46, NewLocationID.Length);
                        Array.Copy(ServiceProviderID, 0, body, 48, ServiceProviderID.Length);
                        Array.Copy(NewServiceProviderID, 0, body, 49, NewServiceProviderID.Length);
                        Array.Copy(MicroPaymentFlag, 0, body, 52, MicroPaymentFlag.Length);
                        Array.Copy(OneDayQuotaForMicroPayment, 0, body, 53, OneDayQuotaForMicroPayment.Length);
                        Array.Copy(OnceQuotaForMicroPayment, 0, body, 55, OnceQuotaForMicroPayment.Length);
                        Array.Copy(SAMSlotControlFlag, 0, body, 57, SAMSlotControlFlag.Length);
                        Array.Copy(MifareKetSet, 0, body, 58, MifareKetSet.Length);
                        body[body.Length - 2] = 0xFA;
                        body[body.Length - 1] = GetLRC(body);
                        break;
                    }
                case 2://尋卡
                    {
                        byte[] SearchCardCMD = { 0x02, 0x01 };
                        Command = SearchCardCMD;
                        Array.Resize(ref body, 1);
                        body[0] = 0x00;
                        break;
                    }
                case 3://遠鑫讀卡
                    {

                        byte[] YHDPReadCMD = { 0x02, 0x0B };
                        byte[] YHDPReadBody = { 0x02, 0x01, 0x00, 0x00, 0x53, 0x07, 0x01, 0x03, 0x04, 0x05, 0x06, 0x08, 0x0A };
                        Command = YHDPReadCMD;
                        Array.Resize(ref body, 14);
                        Array.Copy(YHDPReadBody,0, body,0, YHDPReadBody.Length);
                        body[body.Length-1] = GetLRC(body);
                        break;
                    }
                case 4://悠遊卡讀卡
                    {
                        byte[] PPRResetCMD = { 0x04, 0x01 };
                        Command = PPRResetCMD;
                        Array.Resize(ref body, 23);
                        byte[] CLAToLc = { 0x80, 0x05, 0x01, 0x00, 0x10 };
                        byte[] TMSerialNumber = { 0x30, 0x30, 0x30, 0x30, 0x30, 0x31 };
                        byte[] TXNDateTime = GetUnixTime();
                        byte[] DataType = { 0x01 };
                        Array.Copy(CLAToLc, 0, body, 0, CLAToLc.Length);
                        Array.Copy(TMSerialNumber, 0, body, 6, TMSerialNumber.Length);
                        Array.Copy(TXNDateTime, 0, body, 12, TXNDateTime.Length);
                        Array.Copy(DataType, 0, body, 16, DataType.Length);
                        body[body.Length - 2] = 0xFB;
                        body[body.Length - 1] = GetLRC(body);
                        break;
                    }
                case 5://一卡通讀卡
                    {
                        byte[] ipassReadCMD = { 0x05, 0x01 };
                        Command = ipassReadCMD;
                        Array.Resize(ref body, 9);
                        Array.Copy(CardNumber, 0, body, 0, 4);
                        Array.Copy(GetUnixTime(), 0, body, 4, 4);
                        body[body.Length - 1] = GetLRC(body);
                        break;
                    }
                case 6://愛金卡讀卡
                    {
                        byte[] iCASHReadCMD = { 0x06, 0x01 };
                        Command = iCASHReadCMD;
                        Array.Resize(ref body, 13);
                        body[0] = 0x07;
                        Array.Copy(CardNumber, 0, body, 1, CardNumber.Length);
                        Array.Copy(GetUnixTime(), 0, body, 8, 4);
                        body[body.Length - 1] = GetLRC(body);
                        break;
                    }
            }
            Length = BitConverter.GetBytes(Convert.ToInt16(body.Length));
            Array.Reverse(Length);
            byte[] CMD = new byte[Header.Length + Command.Length + Length.Length + body.Length + Tail.Length];
            Array.Copy(Header, 0, CMD, 0, Header.Length);
            Array.Copy(Command, 0, CMD, 1, Command.Length);
            Array.Copy(Length, 0, CMD, 3, Length.Length);
            Array.Copy(body, 0, CMD, 5, body.Length);
            Array.Copy(Tail, 0, CMD, CMD.Length - 2, Tail.Length);
            Console.WriteLine(BitConverter.ToString(CMD, 0, CMD.Length));
            try
            {
                RS232.Write(CMD, 0, CMD.Length);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            TimeOutTimer.Start();
        }
        private byte GetLRC(byte[] data)
        {
            byte LRC = 0x00;
            for (int i = 0; i < data.Length; i++)
            {
                LRC ^= data[i];
            }
            return LRC;
        }
        private byte[] GetUnixTime()
        {
            byte[] UnixTime;
            UnixTime = BitConverter.GetBytes(Convert.ToUInt32(DateTimeOffset.Now.ToUnixTimeSeconds()));
            return UnixTime;
        }
        /***********************************************
         * DataReceived:資料接收流程                   *
         * 步驟一:確認有無資料     步驟二:資料存入佇列 *
         * 步驟三:檢查資料是否完整                     *
         ***********************************************/
        private void DataReceived()
        {
            int RecievedLength;
            while (ThreadFlag)
            {
                try
                {
                    while (RS232.BytesToRead > 0)
                    {
                        ReceivedQueue.Enqueue(Convert.ToByte(RS232.ReadByte()));
                    }
                    if (ReceivedQueue.Count > 5)
                    {
                        ReceivedArray = ReceivedQueue.ToArray();
                        RecievedLength = ReceivedArray[3] * 256 + ReceivedArray[4];
                        if (ReceivedArray.Length == RecievedLength + 7)
                        {
                            TimeOutTimer.Stop();
                            ReceivedQueue.Clear();
                            ReceivedFlag = true;
                            Console.WriteLine("資料已接收完畢");
                        }
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Thread.Sleep(5);
            }
        }
        /*************************************************************
         * ReaderFunction:讀卡機流程                                 *
         * 步驟一:確認鮑率           步驟二:讀取版本號，確認卡機版本 *
         * 步驟三:悠遊卡PPRREST      步驟四:開始尋卡(250ms/次)       *
         * 步驟五:讀卡並顯示卡片內容                                 *
         *************************************************************/
        private void ReaderFunction()
        {
            int status = 1;
            string Version;
            byte CardCompany = 0x00;

            while (ThreadFlag)
            {
                switch (status)
                {
                    case 1://確認鮑率
                        {
                            int counter = 0;
                            CMDWrite(0);
                            while (ReceivedFlag == false && counter < 2)
                            {
                                if (TimeOutFlag)
                                {
                                    counter++;
                                    TimeOutFlag = false;
                                    Delegate.UpdateUI(RS232.BaudRate + "非此讀卡機鮑率" + Environment.NewLine, Info_RichTextBox);
                                    Console.WriteLine(RS232.BaudRate + "非此讀卡機鮑率");
                                    if (RS232.BaudRate == 57600)
                                    {
                                        RS232.BaudRate = 115200;
                                    }
                                    else
                                    {
                                        RS232.BaudRate = 57600;
                                    }
                                    CMDWrite(0);
                                }
                                Thread.Sleep(5);
                            }
                            if (ReceivedFlag)
                            {
                                ReceivedFlag = false;
                                status = 2;
                                Delegate.UpdateUI("讀卡機鮑率:" + RS232.BaudRate + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("讀卡機鮑率:" + RS232.BaudRate);
                            }
                            else
                            {
                                Delegate.UpdateUI("讀卡機連接失敗，請重新連接" + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("讀卡機連接失敗，請重新連接");
                                CloseReader();
                            }
                            break;
                        }
                    case 2://讀取版本號，確認卡機版本
                        {
                            Version = Encoding.ASCII.GetString(ReceivedArray, 5, 13);
                            Delegate.UpdateUI("卡機版本號:" + Version + Environment.NewLine, Info_RichTextBox);
                            Console.WriteLine("卡機版本號:" + Version);
                            if (Version.Substring(0, 3) == "B1M" || Version.Substring(0, 3) == "F1M")
                            {
                                Delegate.UpdateUI("一代卡機" + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("一代卡機");
                                Array.Clear(ReceivedArray, 0, ReceivedArray.Length);
                                CloseReader();
                            }
                            else if (Version.Substring(0, 3) == "F2D" || Version.Substring(0, 3) == "TS2" || Version.Substring(0, 3) == "FRD")
                            {
                                Delegate.UpdateUI("二代卡機" + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("二代卡機");
                                Array.Clear(ReceivedArray, 0, ReceivedArray.Length);
                                status = 3;
                            }
                            else
                            {
                                Delegate.UpdateUI("無法識別此讀卡機版本" + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("無法識別此讀卡機版本");
                                Array.Clear(ReceivedArray, 0, ReceivedArray.Length);
                                CloseReader();
                            }
                            break;
                        }
                    case 3://悠遊卡PPRREST
                        {
                            CMDWrite(1);
                            while (ReceivedFlag == false && TimeOutFlag == false)
                            {
                                if (ThreadFlag == false)
                                {
                                    break;
                                }
                                Thread.Sleep(10);
                            }
                            if(ReceivedFlag && ReceivedArray.Length==260)
                            {
                                
                                Delegate.UpdateUI("PPRReset成功" + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("PPRReset成功");
                            }
                            else
                            {
                                Delegate.UpdateUI("PPRReset失敗" + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("PPRReset失敗");
                            }
                            status = 4;
                            ReceivedFlag = false;
                            TimeOutFlag = false;
                            Array.Clear(ReceivedArray, 0, ReceivedArray.Length);
                            Delegate.UpdateUI(ComPort_ComboBox, true);
                            break;
                        }
                    case 4://開始尋卡(250ms/次)
                        {
                            CMDWrite(2);
                            while (ReceivedFlag == false && TimeOutFlag == false)
                            {
                                if (ThreadFlag==false)
                                {
                                    break;
                                }
                                Thread.Sleep(10);
                            }
                            if (ReceivedFlag && (ReceivedArray.Length == 8 || ReceivedArray.Length == 19))
                            {
                                switch (ReceivedArray[5])
                                {
                                    case 0x01:  //無卡片
                                        {
                                            Array.Clear(CardNumber, 0, CardNumber.Length);
                                            Console.WriteLine("無卡片");
                                            Thread.Sleep(250);
                                            break;
                                        }
                                    case 0x02:  //多卡重疊
                                        {
                                            Array.Clear(CardNumber, 0, CardNumber.Length);
                                            Console.WriteLine("多卡重疊");
                                            Thread.Sleep(250);
                                            break;
                                        }
                                    case 0x00:  //未知
                                        {
                                            Array.Clear(CardNumber, 0, CardNumber.Length);
                                            Console.WriteLine("無法識別卡片");
                                            Thread.Sleep(250);
                                            break;
                                        }
                                    case 0x03:  //遠鑫卡
                                        {
                                            if (BitConverter.ToString(ReceivedArray, 6, 7) != BitConverter.ToString(CardNumber, 0, 7))
                                            {
                                                if (CardCompany != ReceivedArray[5])
                                                {
                                                    Delegate.UpdateUI(Content_ListView, YHDP.ListViewItem, YHDP.ReadItem);
                                                    CardCompany = ReceivedArray[5];
                                                }
                                                Array.Copy(ReceivedArray, 6, CardNumber, 0, 7);
                                                Delegate.UpdateUI("遠鑫卡:" + BitConverter.ToString(ReceivedArray, 6, 4) + Environment.NewLine, Info_RichTextBox);
                                                Console.WriteLine("遠鑫卡:" + BitConverter.ToString(ReceivedArray, 6, 4));
                                                status = 5;
                                            }
                                            else
                                            {
                                                Console.WriteLine("同卡");
                                                Thread.Sleep(250);
                                            }
                                            break;
                                        }
                                    case 0x04:  //悠遊卡
                                        {
                                            if (BitConverter.ToString(ReceivedArray, 6, 7) != BitConverter.ToString(CardNumber, 0, 7))
                                            {
                                                if (CardCompany != ReceivedArray[5])
                                                {
                                                    Delegate.UpdateUI(Content_ListView, EasyCard.ListViewItem, EasyCard.ReadItem);
                                                    CardCompany = ReceivedArray[5];
                                                }
                                                Array.Copy(ReceivedArray, 6, CardNumber, 0, 7);
                                                Delegate.UpdateUI("悠遊卡:" + BitConverter.ToString(ReceivedArray, 6, 4) + Environment.NewLine, Info_RichTextBox);
                                                Console.WriteLine("悠遊卡:" + BitConverter.ToString(ReceivedArray, 6, 4));
                                                status = 5;
                                            }
                                            else
                                            {
                                                Console.WriteLine("同卡");
                                                Thread.Sleep(250);
                                            }
                                            break;
                                        }
                                    case 0x05:  //一卡通
                                        {
                                            if (BitConverter.ToString(ReceivedArray, 6, 7) != BitConverter.ToString(CardNumber, 0, 7))
                                            {
                                                if (CardCompany != ReceivedArray[5])
                                                {
                                                    Delegate.UpdateUI(Content_ListView, ipass.ListViewItem, ipass.ReadItem);
                                                    CardCompany = ReceivedArray[5];
                                                }
                                                Array.Copy(ReceivedArray, 6, CardNumber, 0, 7);
                                                Delegate.UpdateUI("一卡通:" + BitConverter.ToString(ReceivedArray, 6, 4) + Environment.NewLine, Info_RichTextBox);
                                                Console.WriteLine("一卡通:" + BitConverter.ToString(ReceivedArray, 6, 4));
                                                status = 5;
                                            }
                                            else
                                            {
                                                Console.WriteLine("同卡");
                                                Thread.Sleep(250);
                                            }
                                            break;
                                        }
                                    case 0x06:  //愛金卡
                                        {
                                            if (BitConverter.ToString(ReceivedArray, 6, 7) != BitConverter.ToString(CardNumber, 0, 7))
                                            {
                                                if (CardCompany!=ReceivedArray[5])
                                                {
                                                    Delegate.UpdateUI(Content_ListView,iCASH.ListViewItem,iCASH.ReadItem);
                                                    CardCompany = ReceivedArray[5];
                                                }
                                                Array.Copy(ReceivedArray, 6, CardNumber, 0, 7);
                                                Delegate.UpdateUI("愛金卡:" + BitConverter.ToString(ReceivedArray, 6, 7) + Environment.NewLine, Info_RichTextBox);
                                                Console.WriteLine("愛金卡:" + BitConverter.ToString(ReceivedArray, 6, 7));
                                                status = 5;
                                            }
                                            else
                                            {
                                                Console.WriteLine("同卡");
                                                Thread.Sleep(250);
                                            }
                                            break;
                                        }
                                }
                            }
                            else
                            {
                                Delegate.UpdateUI("尋卡失敗" + Environment.NewLine, Info_RichTextBox);
                                Console.WriteLine("尋卡失敗");
                                Thread.Sleep(250);
                            }
                            ReceivedFlag = false;
                            TimeOutFlag = false;
                            Array.Clear(ReceivedArray, 0, ReceivedArray.Length);
                            break;
                        }
                    case 5://讀卡並顯示卡片內容
                        {
                            switch (CardCompany)
                            {
                                case 0x03:  //遠鑫卡
                                    {
                                        CMDWrite(3);
                                        while (ReceivedFlag == false && TimeOutFlag == false)
                                        {
                                            if (ThreadFlag == false)
                                            {
                                                break;
                                            }
                                            Thread.Sleep(5);
                                        }
                                        if (ReceivedFlag && ReceivedArray.Length == 354)
                                        {
                                            Delegate.UpdateUI(Content_ListView, YHDP.ListViewItem, YHDP.ByteArrayToListViewSubItem(ReceivedArray));
                                        }
                                        else
                                        {
                                            Delegate.UpdateUI("遠鑫卡讀卡失敗" + Environment.NewLine, Info_RichTextBox);
                                            Console.WriteLine("遠鑫卡讀卡失敗");
                                        }
                                        break;
                                    }
                                case 0x04:  //悠遊卡
                                    {
                                        CMDWrite(4);
                                        while (ReceivedFlag == false && TimeOutFlag == false)
                                        {
                                            if (ThreadFlag == false)
                                            {
                                                break;
                                            }
                                            Thread.Sleep(5);
                                        }
                                        if (ReceivedFlag && ReceivedArray.Length == 261)
                                        {
                                            Delegate.UpdateUI(Content_ListView, EasyCard.ListViewItem, EasyCard.ByteArrayToListViewSubItem(ReceivedArray));
                                        }
                                        else
                                        {
                                            Delegate.UpdateUI("悠遊卡讀卡失敗" + Environment.NewLine, Info_RichTextBox);
                                            Console.WriteLine("悠遊卡讀卡失敗");
                                        }
                                        break;
                                    }
                                case 0x05:  //一卡通
                                    {
                                        CMDWrite(5);
                                        while (ReceivedFlag == false && TimeOutFlag == false)
                                        {
                                            if (ThreadFlag == false)
                                            {
                                                break;
                                            }
                                            Thread.Sleep(5);
                                        }
                                        if (ReceivedFlag && ReceivedArray.Length >= 197)
                                        {
                                            Delegate.UpdateUI(Content_ListView, ipass.ListViewItem, ipass.ByteArrayToListViewSubItem(ReceivedArray));
                                        }
                                        else
                                        {
                                            Delegate.UpdateUI("一卡通讀卡失敗" + Environment.NewLine, Info_RichTextBox);
                                            Console.WriteLine("一卡通讀卡失敗");
                                        }
                                        break;
                                    }
                                case 0x06:  //愛金卡
                                    {
                                        CMDWrite(6);
                                        while (ReceivedFlag == false && TimeOutFlag == false)
                                        {
                                            if (ThreadFlag == false)
                                            {
                                                break;
                                            }
                                            Thread.Sleep(5);
                                        }
                                        if(ReceivedFlag && ReceivedArray.Length==133)
                                        {
                                            Delegate.UpdateUI(Content_ListView, iCASH.ListViewItem, iCASH.ByteArrayToListViewSubItem(ReceivedArray));
                                        }
                                        else
                                        {
                                            Delegate.UpdateUI("愛金卡讀卡失敗" + Environment.NewLine, Info_RichTextBox);
                                            Console.WriteLine("愛金卡讀卡失敗");
                                        }
                                        break;
                                    }
                            }
                            ReceivedFlag = false;
                            TimeOutFlag = false;
                            Array.Clear(ReceivedArray, 0, ReceivedArray.Length);
                            status = 4;
                            break;
                        }
                }
            }
        }
    }
}
public class MyDelegate
{
    private Form Form;
    private delegate void UpdateUIText(string Str, Control ctl);
    private delegate void ChangeEnable(ComboBox comboBox, bool enablebool);
    private delegate void ChangeSelectedIndex(ComboBox comboBox);
    private delegate void ClearListViewItem(ListView listview);
    private delegate void UpdateListViewItem(ListView listview, ListViewItem[] listviewitem, string[] readitem);
    private delegate void UpdateListViewSubItems(ListView listview, ListViewItem[] listviewitem, ListViewSubItem[] data);

    public MyDelegate(Form Form)
    {
        this.Form = Form;
    }
    public void UpdateUI(string str, Control ctl)
    {
        try
        {
            if (Form.InvokeRequired)
            {
                UpdateUIText Update = new UpdateUIText(UpdateUI);
                Form.Invoke(Update, str, ctl);
            }
            else
            {
                ctl.Text += str;
            }
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    public void UpdateUI(ComboBox comboBox,bool enablebool)
    {
        try
        {
            if (Form.InvokeRequired)
            {
                ChangeEnable Update = new ChangeEnable(UpdateUI);
                Form.Invoke(Update, comboBox, enablebool);
            }
            else
            {
                comboBox.Enabled = enablebool;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    public void UpdateUI(ComboBox comboBox)
    {
        try
        {
            if (Form.InvokeRequired)
            {
                ChangeSelectedIndex Update = new ChangeSelectedIndex(UpdateUI);
                Form.Invoke(Update, comboBox);
            }
            else
            {
                comboBox.SelectedIndex = 0;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    public void UpdateUI(ListView listview)
    {
        try
        {
            if (Form.InvokeRequired)
            {
                ClearListViewItem Update = new ClearListViewItem(UpdateUI);
                Form.Invoke(Update, listview);
            }
            else
            {
                listview.Items.Clear();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    public void UpdateUI(ListView listview, ListViewItem[] listviewitem, string[] readitem)
    {
        try
        {
            if (Form.InvokeRequired)
            {
                UpdateListViewItem Update = new UpdateListViewItem(UpdateUI);
                Form.Invoke(Update, listview, listviewitem, readitem);
            }
            else
            {
                listview.Items.Clear();
                for (int i = 0; i < listviewitem.Length; i++)
                {
                    listviewitem[i].SubItems.Clear();
                    listviewitem[i].Text = readitem[i];
                }
                listview.Items.AddRange(listviewitem);
                if(listviewitem.Length == 88)
                {
                    listview.Items[40].BackColor = Color.Yellow;
                    listview.Items[54].BackColor = Color.Yellow;
                    listview.Items[63].BackColor = Color.Yellow;
                    listview.Items[69].BackColor = Color.Yellow;
                    listview.Items[75].BackColor = Color.Yellow;
                    listview.Items[81].BackColor = Color.Yellow;
                }
                else if (listviewitem.Length == 189)
                {
                    listview.Items[2].BackColor = Color.Yellow;
                    listview.Items[11].BackColor = Color.Yellow;
                    listview.Items[23].BackColor = Color.Yellow;
                    listview.Items[25].BackColor = Color.Yellow;
                    listview.Items[35].BackColor = Color.Yellow;
                    listview.Items[44].BackColor = Color.Yellow;
                    listview.Items[53].BackColor = Color.Yellow;
                    listview.Items[62].BackColor = Color.Yellow;
                    listview.Items[71].BackColor = Color.Yellow;
                    listview.Items[80].BackColor = Color.Yellow;
                    listview.Items[89].BackColor = Color.Yellow;
                    listview.Items[98].BackColor = Color.Yellow;
                    listview.Items[107].BackColor = Color.Yellow;
                    listview.Items[113].BackColor = Color.Yellow;
                    listview.Items[123].BackColor = Color.Yellow;
                    listview.Items[125].BackColor = Color.Yellow;
                    listview.Items[138].BackColor = Color.Yellow;
                    listview.Items[147].BackColor = Color.Yellow;
                    listview.Items[156].BackColor = Color.Yellow;
                    listview.Items[167].BackColor = Color.Yellow;
                    listview.Items[176].BackColor = Color.Yellow;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    public void UpdateUI(ListView listview, ListViewItem[] listviewitem, ListViewSubItem[] data)
    {
        try
        {
            if (Form.InvokeRequired)
            {
                UpdateListViewSubItems Update = new UpdateListViewSubItems(UpdateUI);
                Form.Invoke(Update, listview, listviewitem, data);
            }
            else
            {
                for (int i = 0; i < data.Length; i++)
                {
                    listview.Items[i].SubItems.Insert(1, data[i]);

                    if (listview.Items[i].SubItems.Count > 2 && listview.Items[i].SubItems[1].Text != "" && listview.Items[i].SubItems[2].Text != "")
                    {
                        if (listview.Items[i].SubItems[1].Text != listview.Items[i].SubItems[2].Text)
                        {
                            listview.Items[i].BackColor = Color.Red;
                        }
                        else
                        {
                            listview.Items[i].BackColor = Color.White;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
}
public class iCASH
{
    public ListViewItem[] ListViewItem;
    public static string[] ReadItem =
        {
            "icasH2.0卡號","票卡卡片額度","票卡交易序號","卡片種類","卡片有效日期",
            "票卡種類","區碼","票卡點數","身分證字號","身分識別碼",
            "發卡單位","啟用日期","身份有效日期","重置日期","使用上限",
            "前次轉乘代碼","本次轉乘代碼","轉乘日期","轉乘交易系統編號","轉乘業者代碼",
            "轉乘交易類別","轉乘優惠金額","轉乘場站代碼","轉乘設備編號","前次段次交易系統編號",
            "前次段碼","前次段次交易時間","前次段次路線編號","前次段次票卡交易序號","前次段次交易金額",
            "前次上下車狀態","前次段次交易類別","前次往返程註記","前次段次交易業者代號","前次段次交易場站代碼",
            "前次段次設備編號","前次里程交易系統編號","前次里程交易時間","前次里程路線編號","前次里程票卡交易序號",
            "前次里程交易金額","前次上下車狀態","前次里程交易類別","前次里程交易模式","前次往返程註記",
            "前次里程交易業者代號","前次里程交易場站代碼","前次里程設備編號","上車站到終點站票價","搭乘次數",
            "搭乘日期"
        };
    /***************************************************************************************************/
    public iCASH()
    {
        ListViewItem = new ListViewItem[ReadItem.Length];
        for (int i=0;i<ListViewItem.Length;i++)
        {
            ListViewItem[i] = new ListViewItem();
        }
    }
    public ListViewSubItem[] ByteArrayToListViewSubItem(byte[] data)
    {
        ListViewSubItem[] ListViewSubItem = new ListViewSubItem[ListViewItem.Length];
        for (int i = 0; i < ListViewItem.Length; i++)
        {
            ListViewSubItem[i] = new ListViewSubItem();
        }
        ListViewSubItem[0].Text = BitConverter.ToString(data, 5, 8).Replace("-","");//icasH2.0卡號 (BCD)
        ListViewSubItem[1].Text = Convert.ToString(BitConverter.ToInt32(data, 13));//票卡卡片額度
        ListViewSubItem[2].Text = Convert.ToString(BitConverter.ToUInt32(data, 17));//票卡交易序號
        ListViewSubItem[3].Text = BitConverter.ToString(data,21,1);//卡片種類
        ListViewSubItem[4].Text =$"{BitConverter.ToString(data, 22, 2).Replace("-", "")}年{BitConverter.ToString(data, 24, 1)}月{BitConverter.ToString(data, 25, 1)}日";//卡片有效日期
        ListViewSubItem[5].Text = BitConverter.ToString(data, 26, 1);//票卡種類
        ListViewSubItem[6].Text = BitConverter.ToString(data, 27, 1);//區碼
        ListViewSubItem[7].Text = Convert.ToString(BitConverter.ToUInt16(data, 28));//票卡點數(敬老、愛陪)
        ListViewSubItem[8].Text = BitConverter.ToString(data, 30, 10).Replace("-", "");//身分證字號
        ListViewSubItem[9].Text = BitConverter.ToString(data, 40, 1);//身分識別碼
        ListViewSubItem[10].Text = BitConverter.ToString(data, 41, 1);//發卡單位
        ListViewSubItem[11].Text = ByteArrayToUnixTimeString(data, 42);//啟用日期
        ListViewSubItem[12].Text = ByteArrayToUnixTimeString(data, 46);//身份有效日期
        ListViewSubItem[13].Text = ByteArrayToUnixTimeString(data, 50);//重置日期
        ListViewSubItem[14].Text = BitConverter.ToString(data, 54, 4).Replace("-", "");//使用上限 (BCD)
        ListViewSubItem[15].Text = BitConverter.ToString(data, 58, 1);//前次轉乘代碼
        ListViewSubItem[16].Text = BitConverter.ToString(data, 59, 1);//本次轉乘代碼
        ListViewSubItem[17].Text = ByteArrayToUnixTimeString(data, 60);//轉乘日期
        ListViewSubItem[18].Text = BitConverter.ToString(data, 64, 1);//轉乘交易系統編號
        ListViewSubItem[19].Text = BitConverter.ToString(data, 65, 1);//轉乘業者代碼
        ListViewSubItem[20].Text = BitConverter.ToString(data, 66, 1);//轉乘交易類別
        ListViewSubItem[21].Text = Convert.ToString(BitConverter.ToUInt16(data, 67));//轉乘優惠金額
        ListViewSubItem[22].Text = Convert.ToString(BitConverter.ToUInt16(data, 69));//轉乘場站代碼
        ListViewSubItem[23].Text = Convert.ToString(BitConverter.ToUInt32(data, 71));//轉乘設備編號
        /***********************************************************************************/
        ListViewSubItem[24].Text = BitConverter.ToString(data, 75, 1);//前次段次交易系統編號
        ListViewSubItem[25].Text = BitConverter.ToString(data, 76, 1);//前次段碼
        ListViewSubItem[26].Text = ByteArrayToUnixTimeString(data, 77);//前次段次交易時間
        ListViewSubItem[27].Text = Convert.ToString(BitConverter.ToUInt16(data, 81));//前次段次路線編號
        ListViewSubItem[28].Text = Convert.ToString(BitConverter.ToUInt32(data, 83));//前次段次票卡交易序號
        ListViewSubItem[29].Text = Convert.ToString(BitConverter.ToUInt16(data, 87));//前次段次交易金額
        ListViewSubItem[30].Text = BitConverter.ToString(data, 89, 1);//前次上下車狀態(0x01 -> 上車,0x00 -> 下車)
        ListViewSubItem[31].Text = BitConverter.ToString(data, 90, 1);//前次段次交易類別
        ListViewSubItem[32].Text = BitConverter.ToString(data, 91, 1);//前次往返程註記(0x01 -> 去程,0x02 -> 返程,0x00 -> 循環)
        ListViewSubItem[33].Text = BitConverter.ToString(data, 92, 1);//前次段次交易業者代號
        ListViewSubItem[34].Text = Convert.ToString(BitConverter.ToUInt16(data, 93));//前次段次交易場站代碼
        ListViewSubItem[35].Text = Convert.ToString(BitConverter.ToUInt32(data, 95));//前次段次設備編號
        /***********************************************************************************/
        ListViewSubItem[36].Text = BitConverter.ToString(data, 99, 1);//前次里程交易系統編號
        ListViewSubItem[37].Text = ByteArrayToUnixTimeString(data, 100);//前次里程交易時間
        ListViewSubItem[38].Text = Convert.ToString(BitConverter.ToUInt16(data, 104));//前次里程路線編號
        ListViewSubItem[39].Text = Convert.ToString(BitConverter.ToUInt32(data, 106));//前次里程票卡交易序號
        ListViewSubItem[40].Text = Convert.ToString(BitConverter.ToUInt16(data, 110));//前次里程交易金額
        ListViewSubItem[41].Text = BitConverter.ToString(data, 112, 1);//前次上下車狀態(0x01 -> 上車,0x00 -> 下車)
        ListViewSubItem[42].Text = BitConverter.ToString(data, 113, 1);//前次里程交易類別
        ListViewSubItem[43].Text = BitConverter.ToString(data, 114, 1);//前次里程交易模式
        ListViewSubItem[44].Text = BitConverter.ToString(data, 115, 1);//前次往返程註記(0x01 -> 去程,0x02 -> 返程,0x00 -> 循環)
        ListViewSubItem[45].Text = BitConverter.ToString(data, 116, 1);//前次里程交易業者代號
        ListViewSubItem[46].Text = Convert.ToString(BitConverter.ToUInt16(data, 117));//前次里程交易場站代碼
        ListViewSubItem[47].Text = Convert.ToString(BitConverter.ToUInt32(data, 119));//前次里程設備編號
        ListViewSubItem[48].Text = Convert.ToString(BitConverter.ToUInt16(data, 123));//上車站到終點站票價
        ListViewSubItem[49].Text = BitConverter.ToString(data, 125, 1);//搭乘次數
        ListViewSubItem[50].Text = Convert.ToString(BitConverter.ToUInt32(data, 126));//搭乘日期
        return ListViewSubItem;
    }
    private string ByteArrayToUnixTimeString(byte[] data, int startindex)
    {
        string UnixTimeString;
        if (BitConverter.ToUInt32(data, startindex) > 0)
        {
            UnixTimeString = new DateTime(1970, 1, 1, 0, 0, 0).AddSeconds(BitConverter.ToInt32(data, startindex)).ToString();
        }
        else
        {
            UnixTimeString = "無";
        }
        return UnixTimeString;
    }
}
public class YHDP
{
    public ListViewItem[] ListViewItem;
    public static string[] ReadItem =
        {
            "錢包金額","錢包狀態",
            "發行管理資料",
            "發卡單位編號","發卡設備編號","發行批號","發出日期","有效日期","卡片格式版本","卡片狀態","檢查碼",
            "票值管理資料",
            "自動加值設定","自動加值票值數額","儲存最大票值數額","每筆可扣減最大票值數額","指定加值設定","指定加值票值數額","自動加值日期","連續離線自動加值次數","連續自動加值次數","連續指定加值次數","檢查碼",
            "卡片防偽驗證資料",
            "防偽驗證資料",
            "卡片交易狀態資料",
            "卡片交易序號","交易紀錄指標","優惠積點數","優惠積點交易序號","鎖卡旗標","進出閘門口編號","進出閘門口時間","轉乘Flag(交易類別)","轉乘Flag(交易群組)",
            "最近兩筆匣門交易紀錄(1)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "最近兩筆匣門交易紀錄(2)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "最近六筆交易紀錄(1)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "最近六筆交易紀錄(2)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "最近六筆交易紀錄(3)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "最近六筆交易紀錄(4)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "最近六筆交易紀錄(5)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "最近六筆交易紀錄(6)",
            "交易序號","交易時間","交易類別","交易票值","交易後票值","交易系統編號","交易地點/運輸業者","交易機器",
            "卡片管理1",
            "卡種","使用者截止日期","使用者序號","發卡企業編號","卡片押金",
            "卡片管理2",
            "卡別","記名註記","生日","產品別","特種票類別代碼","特種票終止日期","特種票最大可使用次數/點數/金額","特種票已使用次數/點數/金額","特種票有效天數",
            "卡片管理3",
            "儲值檔識別碼",
            "里程客運進出站交易管理資料",
            "客運公司編號","最後一次上/下車碼","里程特種票類別代碼","里程特種票原始可用次數/額度","里程特種票剩餘可用次數/額度","里程特種票有效日期","里程特種票有效天數",
            "里程特種票有效起站","里程特種票有效迄站","當日累積里程交易日期","當日累積里程搭乘金額","最後一次搭乘路線編號",
            "里程計費上車交易紀錄",
            "交易時間","交易票值","交易後票值","交易類別","上車站別ID","交易序號","交易地點/運輸業者","行駛方向",
            "里程計費下車交易紀錄",
            "交易時間","交易票值","交易後票值","交易類別","上車站別ID","交易序號","交易地點/運輸業者","行駛方向",
            "鐵路重要資料備份區",
            "最後一次進/出站編號","最後交易狀態","最後交易時間","台鐵特種票類別代碼","台鐵特種票有效日期","台鐵特種票有效天數","台鐵特種票有效起站","台鐵特種票有效迄站","台鐵特種票原始可用次數","台鐵特種票剩餘可用次數",
            "里程客運重要資料備份區",
            "交易時間","交易票值","交易後票值","交易類別","上車站別ID","交易序號","交易地點/運輸業者","行駛方向",
            "里程客運特種票資料區",
            "客運公司編號","最後一次上/下車碼","里程特種票類別代碼","里程特種票原始可用次數/額度","里程特種票剩餘可用次數/額度","里程特種票有效日期",
            "里程特種票有效天數","里程特種票有效起站","里程特種票有效迄站","當日累積里程交易日期","當日累積里程搭乘金額","最後一次搭乘路線編號"
        };
    /***************************************************************************************************/
    public YHDP()
    {
        ListViewItem = new ListViewItem[ReadItem.Length];
        for (int i = 0; i < ListViewItem.Length; i++)
        {
            ListViewItem[i] = new ListViewItem();
        }
    }
    public ListViewSubItem[] ByteArrayToListViewSubItem(byte[] data)
    {
        int index;
        ListViewSubItem[] ListViewSubItem = new ListViewSubItem[ListViewItem.Length];
        for (int i = 0; i < ListViewItem.Length; i++)
        {
            ListViewSubItem[i] = new ListViewSubItem();
        }
        ListViewSubItem[0].Text = Convert.ToString(MSBToLSB(data, 5,4,false));//錢包金額
        ListViewSubItem[1].Text = BitConverter.ToString(data, 9, 1);//錢包狀態
        //發行管理資料
        ListViewSubItem[3].Text = BitConverter.ToString(data, 15, 1);//發卡單位編號
        ListViewSubItem[4].Text = Convert.ToString(BitConverter.ToUInt16(data, 16));//發卡設備編號
        ListViewSubItem[5].Text = Convert.ToString(MSBToLSB(data, 18, 2, true));//發行批號
        ListViewSubItem[6].Text = UInt32ToUnixTimeString((uint)MSBToLSB(data, 20, 4, true));//發出日期
        ListViewSubItem[7].Text = UInt32ToUnixTimeString((uint)MSBToLSB(data, 24, 4, true));//有效日期
        ListViewSubItem[8].Text = BitConverter.ToString(data, 28, 1);//卡片格式版本
        ListViewSubItem[9].Text = BitConverter.ToString(data, 29, 1);//卡片狀態
        ListViewSubItem[10].Text = BitConverter.ToString(data, 30, 1);//卡片狀態檢查碼
        //票值管理資料
        ListViewSubItem[12].Text = BitConverter.ToString(data, 31, 1);//自動加值設定
        ListViewSubItem[13].Text = Convert.ToString(MSBToLSB(data, 32, 2, true));//自動加值票值數額
        ListViewSubItem[14].Text = Convert.ToString(MSBToLSB(data, 34, 2, true));//儲存最大票值數額
        ListViewSubItem[15].Text = Convert.ToString(MSBToLSB(data, 36, 2, true));//每筆可扣減最大票值數額
        ListViewSubItem[16].Text = BitConverter.ToString(data, 38, 1);//指定加值設定
        ListViewSubItem[17].Text = Convert.ToString(BitConverter.ToUInt16(data, 39));//指定加值票值數額
        ListViewSubItem[18].Text = BitConverter.ToString(data, 41, 2);//自動加值日期
        ListViewSubItem[19].Text = BitConverter.ToString(data, 43, 1);//連續離線自動加值次數
        ListViewSubItem[20].Text = BitConverter.ToString(data, 44, 1);//連續自動加值次數
        ListViewSubItem[21].Text = BitConverter.ToString(data, 45, 1);//連續指定加值次數
        ListViewSubItem[22].Text = BitConverter.ToString(data, 46, 1);//檢查碼
        //卡片防偽驗證資料
        ListViewSubItem[24].Text = BitConverter.ToString(data, 47, 16);//防偽驗證資料
        //卡片交易狀態資料
        ListViewSubItem[26].Text = BitConverter.ToString(BitConverter.GetBytes((ushort)MSBToLSB(data, 63, 2, true)));//卡片交易序號
        ListViewSubItem[27].Text = BitConverter.ToString(data, 65, 1);//交易紀錄指標
        ListViewSubItem[28].Text = Convert.ToString(BitConverter.ToUInt16(data, 66));//優惠積點數
        ListViewSubItem[29].Text = Convert.ToString(BitConverter.ToUInt16(data, 68));//優惠積點交易序號
        ListViewSubItem[30].Text = BitConverter.ToString(data, 70, 1);//鎖卡旗標
        ListViewSubItem[31].Text = Convert.ToString(BitConverter.ToUInt16(data, 71));//進出閘門口編號
        ListViewSubItem[32].Text = UInt32ToUnixTimeString(BitConverter.ToUInt32(data, 73));//進出閘門口時間
        ListViewSubItem[33].Text = BitConverter.ToString(data, 77, 1);//轉乘Flag(交易類別)
        ListViewSubItem[34].Text = BitConverter.ToString(data, 78, 1);//轉乘Flag(交易群組)
        //最近兩筆匣門交易紀錄(1~2)
        //最近六筆交易紀錄(1~6)
        index = 35;
        for (int i=0;i<8;i++)
        {
            ListViewSubItem[index + 1 + (8 * i)].Text = BitConverter.ToString(data, 79+i*16, 1);//交易序號
            ListViewSubItem[index + 2 + (8 * i)].Text = UInt32ToUnixTimeString(BitConverter.ToUInt32(data, 80 + i * 16));//交易時間
            ListViewSubItem[index + 3 + (8 * i)].Text = BitConverter.ToString(data, 84 + i * 16, 1);//交易類別
            ListViewSubItem[index + 4 + (8 * i)].Text = Convert.ToString(MSBToLSB(data, 85 + i * 16, 2, true));//交易票值
            ListViewSubItem[index + 5 + (8 * i)].Text = Convert.ToString(BitConverter.ToInt16(data, 87 + i * 16));//交易後票值
            ListViewSubItem[index + 6 + (8 * i)].Text = BitConverter.ToString(data, 89 + i * 16, 1);//交易系統編號
            ListViewSubItem[index + 7 + (8 * i)].Text = BitConverter.ToString(data, 90 + i * 16, 1);//交易地點/運輸業者
            ListViewSubItem[index + 8 + (8 * i)].Text = Convert.ToString(BitConverter.ToUInt32(data, 91 + i * 16));//交易機器
            index++;
        }
        //卡片管理1
        ListViewSubItem[108].Text = BitConverter.ToString(data, 207, 1);//卡種
        ListViewSubItem[109].Text = UInt32ToUnixTimeString((uint)MSBToLSB(data, 208, 4, true));//使用者截止日期
        ListViewSubItem[110].Text = BitConverter.ToString(data, 212, 6);//使用者序號
        ListViewSubItem[111].Text = BitConverter.ToString(data, 218, 1);//發卡企業編號
        ListViewSubItem[112].Text = Convert.ToString(BitConverter.ToUInt16(data, 221));//卡片押金
        //卡片管理2
        ListViewSubItem[114].Text = BitConverter.ToString(data, 223, 1);//卡別
        ListViewSubItem[115].Text = BitConverter.ToString(data, 224, 1);//記名註記
        ListViewSubItem[116].Text = BitConverter.ToString(data, 225, 2);//生日
        ListViewSubItem[117].Text = BitConverter.ToString(data, 227, 1);//產品別
        ListViewSubItem[118].Text = BitConverter.ToString(data, 228, 1);//特種票類別代碼
        ListViewSubItem[119].Text = UInt32ToUnixTimeString(BitConverter.ToUInt32(data, 229));//特種票終止日期
        ListViewSubItem[120].Text = Convert.ToString(BitConverter.ToInt16(data, 233));//特種票最大可使用次數/點數/金額
        ListViewSubItem[121].Text = Convert.ToString(BitConverter.ToInt16(data, 235));//特種票已使用次數/點數/金額
        ListViewSubItem[122].Text = Convert.ToString(BitConverter.ToUInt16(data, 237));//特種票有效天數
        //卡片管理3
        ListViewSubItem[124].Text = BitConverter.ToString(data, 239, 8);//儲值檔識別碼
        //里程客運進出站交易管理資料
        ListViewSubItem[126].Text = BitConverter.ToString(data, 255, 1);//客運公司編號
        ListViewSubItem[127].Text = BitConverter.ToString(data, 256, 1);//最後一次上/下車碼
        ListViewSubItem[128].Text = BitConverter.ToString(data, 256, 1);//里程特種票類別代碼
        ListViewSubItem[129].Text = Convert.ToString(BitConverter.ToInt16(data, 257));//里程特種票原始可用次數/額度
        ListViewSubItem[130].Text = Convert.ToString(BitConverter.ToInt16(data, 259));//里程特種票剩餘可用次數/額度
        ListViewSubItem[131].Text = ByteArrayToDosDateString(data,261);//里程特種票有效日期
        ListViewSubItem[132].Text = data[263].ToString();//里程特種票有效天數
        ListViewSubItem[133].Text = data[264].ToString();//里程特種票有效起站
        ListViewSubItem[134].Text = data[265].ToString();//里程特種票有效迄站
        ListViewSubItem[135].Text = BitConverter.ToString(data, 266, 2);//當日累積里程交易日期
        ListViewSubItem[136].Text = data[268].ToString();//當日累積里程搭乘金額
        ListViewSubItem[137].Text = BitConverter.ToString(data, 269, 2);//最後一次搭乘路線編號
        //里程計費上、下車交易紀錄
        index = 138;
        for (int i=0;i<2;i++)
        {
            ListViewSubItem[index + 1 + (8 * i)].Text = UInt32ToUnixTimeString(BitConverter.ToUInt32(data, 271 + i * 16));//交易時間
            ListViewSubItem[index + 2 + (8 * i)].Text = Convert.ToString(BitConverter.ToInt16(data, 275 + i * 16));//交易票值
            ListViewSubItem[index + 3 + (8 * i)].Text = Convert.ToString(BitConverter.ToInt16(data, 277 + i * 16));//交易後票值
            ListViewSubItem[index + 4 + (8 * i)].Text = BitConverter.ToString(data, 279 + i * 16, 1);//交易類別
            ListViewSubItem[index + 5 + (8 * i)].Text = BitConverter.ToString(data, 280 + i * 16, 1);//上車站別ID
            ListViewSubItem[index + 6 + (8 * i)].Text = BitConverter.ToString(data, 281 + i * 16, 1);//交易序號
            ListViewSubItem[index + 7 + (8 * i)].Text = BitConverter.ToString(BitConverter.GetBytes((uint)MSBToLSB(data, 282 + i * 16, 4, true))).Replace("-","");//交易地點/運輸業者
            ListViewSubItem[index + 8 + (8 * i)].Text = BitConverter.ToString(data, 286 + i * 16, 1);//行駛方向
            index++;
        }
        
        //鐵路重要資料備份區
        ListViewSubItem[157].Text = BitConverter.ToString(data, 303 , 2);//最後一次進/出站編號
        ListViewSubItem[158].Text = BitConverter.ToString(data, 305, 1);//最後交易狀態
        ListViewSubItem[159].Text = UInt32ToUnixTimeString(BitConverter.ToUInt32(data, 306));//最後交易時間
        ListViewSubItem[160].Text = BitConverter.ToString(data, 310, 1);//台鐵特種票類別代碼
        ListViewSubItem[161].Text = ByteArrayToDosDateString(data, 311); ;//台鐵特種票有效日期
        ListViewSubItem[162].Text = data[313].ToString();//台鐵特種票有效天數
        ListViewSubItem[163].Text = BitConverter.ToString(data, 314, 1);//台鐵特種票有效起站
        ListViewSubItem[164].Text = BitConverter.ToString(data, 315, 1);//台鐵特種票有效迄站
        ListViewSubItem[165].Text = data[316].ToString();//台鐵特種票原始可用次數
        ListViewSubItem[166].Text = data[317].ToString();//台鐵特種票剩餘可用次數
        //里程客運重要資料備份區
        ListViewSubItem[168].Text = UInt32ToUnixTimeString(BitConverter.ToUInt32(data, 319));//交易時間
        ListViewSubItem[169].Text = Convert.ToString(BitConverter.ToInt16(data, 323));//交易票值
        ListViewSubItem[170].Text = Convert.ToString(BitConverter.ToInt16(data, 325));//交易後票值
        ListViewSubItem[171].Text = BitConverter.ToString(data, 327, 1);//交易類別
        ListViewSubItem[172].Text = BitConverter.ToString(data, 328, 1);//上車站別ID
        ListViewSubItem[173].Text = BitConverter.ToString(data, 329, 1);//交易序號
        ListViewSubItem[174].Text = BitConverter.ToString(BitConverter.GetBytes((uint)MSBToLSB(data, 330, 4, true))).Replace("-", "");//交易地點/運輸業者
        ListViewSubItem[175].Text = BitConverter.ToString(data, 334, 1);//行駛方向
        //里程客運特種票資料區
        ListViewSubItem[177].Text = BitConverter.ToString(data, 335, 1);//客運公司編號
        ListViewSubItem[178].Text = BitConverter.ToString(data, 336, 1);//最後一次上/下車碼
        ListViewSubItem[179].Text = BitConverter.ToString(data, 336, 1);//里程特種票類別代碼
        ListViewSubItem[180].Text = Convert.ToString(BitConverter.ToInt16(data, 337));//里程特種票原始可用次數/額度
        ListViewSubItem[181].Text = Convert.ToString(BitConverter.ToInt16(data, 339));//里程特種票剩餘可用次數/額度
        ListViewSubItem[182].Text = ByteArrayToDosDateString(data, 341); ;//里程特種票有效日期
        ListViewSubItem[183].Text = data[343].ToString();//里程特種票有效天數
        ListViewSubItem[184].Text = data[344].ToString();//里程特種票有效起站
        ListViewSubItem[185].Text = data[345].ToString();//里程特種票有效迄站
        ListViewSubItem[186].Text = BitConverter.ToString(data, 346, 2);//當日累積里程交易日期
        ListViewSubItem[187].Text = data[348].ToString();//當日累積里程搭乘金額
        ListViewSubItem[188].Text = BitConverter.ToString(data, 349, 2);//最後一次搭乘路線編號
        return ListViewSubItem;
    }
    private object MSBToLSB(byte[] data, int startindex, int length, bool unsigned)
    {
        byte[] LSBArray = { };
        switch (length)
        {
            case 2:
                {
                    Array.Resize(ref LSBArray, 2);
                    Array.Copy(data, startindex, LSBArray, 0, 2);
                    Array.Reverse(LSBArray);
                    if (unsigned)
                    {
                        return BitConverter.ToUInt16(LSBArray, 0);
                    }
                    else
                    {
                        return BitConverter.ToInt16(LSBArray, 0);
                    }
                }
            case 4:
                {
                    Array.Resize(ref LSBArray, 4);
                    Array.Copy(data, startindex, LSBArray, 0, 4);
                    Array.Reverse(LSBArray);
                    if (unsigned)
                    {
                        return BitConverter.ToUInt32(LSBArray, 0);
                    }
                    else
                    {
                        return BitConverter.ToInt32(LSBArray, 0);
                    }
                }
            default:
                {
                    return 0;
                }
        }
    }
    private string UInt32ToUnixTimeString(uint data)
    {
        string UnixTimeString;
        if (data > 0 && data!= 4294967295)
        {
            UnixTimeString = new DateTime(1970,1,1,0,0,0).AddSeconds(data).ToString();
        }
        else
        {
            UnixTimeString = "無";
        }
        return UnixTimeString;
    }
    private string ByteArrayToDosDateString(byte[] data, int startindex)
    {
        byte[] DosDateArray = new byte[2];
        ushort DosDateUShort;
        string DosDateString;
        Array.Copy(data, startindex, DosDateArray, 0, 2);
        Array.Reverse(DosDateArray);
        DosDateUShort = BitConverter.ToUInt16(DosDateArray, 0);
        if (DosDateUShort > 0)
        {
            DosDateString = new DateTime(1980, (DosDateUShort & 0x01E0) >> 5, DosDateUShort & 0x001F).AddYears((DosDateUShort >> 9)).ToString("yyyy/MM/dd");
        }
        else
        {
            DosDateString = "無";
        }
        return DosDateString;
    }
}
public class ipass
{
    public ListViewItem[] ListViewItem;
    public static string[] ReadItem =
        {
            "卡號","主要電子票值","備份電子票值","同步後電子票值","身份證字號",
            "公車端票卡種類","特種票識別身分","特種票識別單位","特種票識別起始日","特種票識別有效日",
            "特種票重置日期","特種票起站代碼","特種票迄站代碼","特種票路線編號","特種票使用限次",
            "公車業者代碼","上次公車端區碼","上次公車端交易運輸業者","上次公車端交易路線","上次公車端交易站號",
            "上次公車端交易時間","上次公車端搭乘狀態","上次搭乘類型","上次公車端交易驗票機編號","特種票已使用次數",
            "卡片交易序號","上次交易時間","上次交易類別","上次交易票值/票點","上次交易後票值/票點",
            "上次交易系統編號","上次交易地點編號","上次交易機器編號","計程預收金額","同步狀態",
            "轉乘ID","轉換旗標","旅遊卡天數","旅遊卡有效日","個人化有效日",
            "#轉乘識別群組",
            "#1交易時間","#2本次系統代碼","#3前次系統代碼","#4交易類別","#5交易業者編號","#6交易地點編號","#7交易路線編號","#8交易設備編號",
            "記名旗標","卡片狀態","記名學生卡識別","定期票業者代碼","桃園市民",
            "MaaS",
            "MaaS卡種類","MaaS","區域代碼","交通運具","票種旗標","天數/時數","起始時間","結束時間",
            "起迄站資訊1",
            "起訖站交通運具","起站","訖站","路線編號","使用次數or累積使用金額",
            "起迄站資訊2",
            "起訖站交通運具","起站","訖站","路線編號","使用次數or累積使用金額",
            "起迄站資訊3",
            "起訖站交通運具","起站","訖站","路線編號","使用次數or累積使用金額",
            "起迄站資訊4",
            "起訖站交通運具","起站","訖站","路線編號","使用次數or累積使用金額",
            "外觀卡號"
        };
    /***************************************************************************************************/
    public ipass()
    {
        ListViewItem = new ListViewItem[ReadItem.Length];
        for (int i = 0; i < ListViewItem.Length; i++)
        {
            ListViewItem[i] = new ListViewItem();
        }
    }
    public ListViewSubItem[] ByteArrayToListViewSubItem(byte[] data)
    {
        ListViewSubItem[] ListViewSubItem = new ListViewSubItem[ListViewItem.Length];
        for (int i = 0; i < ListViewItem.Length; i++)
        {
            ListViewSubItem[i] = new ListViewSubItem();
        }
        ListViewSubItem[0].Text = BitConverter.ToString(data, 5, 16);//卡號
        ListViewSubItem[1].Text = Convert.ToString(BitConverter.ToInt32(data, 21));//主要電子票值
        ListViewSubItem[2].Text = Convert.ToString(BitConverter.ToInt32(data, 25));//備份電子票值
        ListViewSubItem[3].Text = Convert.ToString(BitConverter.ToInt32(data, 29));//同步後電子票值
        ListViewSubItem[4].Text = BitConverter.ToString(data, 33, 6);//身份證字號
        ListViewSubItem[5].Text = BitConverter.ToString(data, 39, 1);//公車端票卡種類
        ListViewSubItem[6].Text = BitConverter.ToString(data, 40, 1);//特種票識別身分
        ListViewSubItem[7].Text = BitConverter.ToString(data, 41, 1);//特種票識別單位
        ListViewSubItem[8].Text = $"{Convert.ToString(BitConverter.ToUInt16(data, 42))}年{data[44]}月{data[45]}日";//特種票識別起始日
        ListViewSubItem[9].Text = $"{Convert.ToString(BitConverter.ToUInt16(data, 46))}年{data[48]}月{data[49]}日";//特種票識別有效日
        ListViewSubItem[10].Text = $"{Convert.ToString(BitConverter.ToUInt16(data, 50))}年{data[52]}月{data[53]}日";//特種票重置日期
        ListViewSubItem[11].Text = BitConverter.ToString(data, 54, 1);//特種票起站代碼
        ListViewSubItem[12].Text = BitConverter.ToString(data, 55, 1);//特種票迄站代碼
        ListViewSubItem[13].Text = Convert.ToString(BitConverter.ToUInt16(data, 56));//特種票路線編號
        ListViewSubItem[14].Text = Convert.ToString(BitConverter.ToUInt16(data, 58));//特種票使用限次
        ListViewSubItem[15].Text = BitConverter.ToString(data, 60, 1);//公車業者代碼
        ListViewSubItem[16].Text = BitConverter.ToString(data, 61, 1);//上次公車端區碼
        ListViewSubItem[17].Text = BitConverter.ToString(data, 62, 1);//上次公車端交易運輸業者
        ListViewSubItem[18].Text = Convert.ToString(BitConverter.ToUInt16(data, 63));//上次公車端交易路線
        ListViewSubItem[19].Text = BitConverter.ToString(data, 65, 1);//上次公車端交易站號
        ListViewSubItem[20].Text = ByteArrayToUnixTimeString(data, 66);//上次公車端交易時間
        ListViewSubItem[21].Text = BitConverter.ToString(data, 70, 1);//上次公車端搭乘狀態
        ListViewSubItem[22].Text = BitConverter.ToString(data, 71, 1);//上次搭乘類型
        ListViewSubItem[23].Text = Convert.ToString(BitConverter.ToUInt16(data, 72));//上次公車端交易驗票機編號
        ListViewSubItem[24].Text = Convert.ToString(BitConverter.ToUInt16(data, 74));//特種票已使用次數
        ListViewSubItem[25].Text = Convert.ToString(BitConverter.ToUInt16(data, 76));//卡片交易序號
        ListViewSubItem[26].Text = ByteArrayToUnixTimeString(data, 78);//上次交易時間
        ListViewSubItem[27].Text = BitConverter.ToString(data, 82, 1);//上次交易類別
        ListViewSubItem[28].Text = Convert.ToString(BitConverter.ToUInt16(data, 83));//上次交易票值/票點
        ListViewSubItem[29].Text = Convert.ToString(BitConverter.ToInt16(data, 85));//上次交易後票值/票點
        ListViewSubItem[30].Text = BitConverter.ToString(data, 87, 1);//上次交易系統編號
        ListViewSubItem[31].Text = BitConverter.ToString(data, 88, 1);//上次交易地點編號
        ListViewSubItem[32].Text = Convert.ToString(BitConverter.ToUInt32(data, 89));//上次交易機器編號
        ListViewSubItem[33].Text = Convert.ToString(data[93]);//計程預收金額
        ListViewSubItem[34].Text = BitConverter.ToString(data, 94, 1);//同步狀態
        ListViewSubItem[35].Text = BitConverter.ToString(data, 95, 2);//轉乘ID
        ListViewSubItem[36].Text = BitConverter.ToString(data, 97, 1);//轉換旗標
        ListViewSubItem[37].Text = Convert.ToString(data[98]);//旅遊卡天數
        ListViewSubItem[38].Text = ByteArrayToUnixTimeString(data, 99);//旅遊卡有效日
        ListViewSubItem[39].Text = ByteArrayToUnixTimeString(data, 103);//個人化有效日
        //轉乘識別群組
        ListViewSubItem[41].Text = ByteArrayToUnixTimeString(data, 107);//交易時間
        ListViewSubItem[42].Text = BitConverter.ToString(data, 111, 1);//本次系統代碼
        ListViewSubItem[43].Text = BitConverter.ToString(data, 112, 1);//前次系統代碼
        ListViewSubItem[44].Text = BitConverter.ToString(data, 113, 1);//交易類別
        ListViewSubItem[45].Text = BitConverter.ToString(data, 114, 1);//交易業者編號
        ListViewSubItem[46].Text = BitConverter.ToString(data, 115, 1);//交易地點編號
        ListViewSubItem[47].Text = Convert.ToString(BitConverter.ToUInt16(data, 116));//交易路線編號
        ListViewSubItem[48].Text = Convert.ToString(BitConverter.ToUInt16(data, 118));//交易設備編號

        ListViewSubItem[49].Text = BitConverter.ToString(data, 120, 1);//記名旗標
        ListViewSubItem[50].Text = BitConverter.ToString(data, 121, 1);//卡片狀態
        ListViewSubItem[51].Text = BitConverter.ToString(data, 122, 1);//記名學生卡識別
        ListViewSubItem[52].Text = BitConverter.ToString(data, 123, 1);//定期票業者代碼
        ListViewSubItem[53].Text = BitConverter.ToString(data, 124, 1);//桃園市民
        //MaaS
        ListViewSubItem[55].Text = BitConverter.ToString(data, 125, 1);//MaaS卡種類
        ListViewSubItem[56].Text = BitConverter.ToString(data, 126, 1);//MaaS
        ListViewSubItem[57].Text = BitConverter.ToString(data, 127, 1);//區域代碼
        ListViewSubItem[58].Text = Convert.ToString(BitConverter.ToUInt32(data, 128));//交通運具
        ListViewSubItem[59].Text = BitConverter.ToString(data, 132, 1);//票種旗標
        ListViewSubItem[60].Text = Convert.ToString(data[133]);//天數/時數
        ListViewSubItem[61].Text = ByteArrayToUnixTimeString(data, 134);//起始時間
        ListViewSubItem[62].Text = ByteArrayToUnixTimeString(data, 138);//結束時間
        //起迄站資訊1~4
        for(int i=0;i<4;i++)
        {
            ListViewSubItem[64 + i * 6].Text = Convert.ToString(data[142 + i * 9]);//起訖站交通運具
            ListViewSubItem[65 + i * 6].Text = Convert.ToString(BitConverter.ToUInt16(data, 143 + i * 9));//起站
            ListViewSubItem[66 + i * 6].Text = Convert.ToString(BitConverter.ToUInt16(data, 145 + i * 9));//訖站
            ListViewSubItem[67 + i * 6].Text = Convert.ToString(BitConverter.ToUInt16(data, 147 + i * 9));//路線編號
            ListViewSubItem[68 + i * 6].Text = Convert.ToString(BitConverter.ToUInt16(data, 149 + i * 9));//使用次數 or 累積使用金額
        }
        ListViewSubItem[87].Text = BitConverter.ToString(data, 178, 16);//外觀卡號
        return ListViewSubItem;
    }
    private string ByteArrayToUnixTimeString(byte[] data, int startindex)
    {
        string UnixTimeString;
        if (BitConverter.ToUInt32(data, startindex) > 0)
        {
            UnixTimeString = new DateTime(1970, 1, 1, 8, 0, 0).AddSeconds(BitConverter.ToInt32(data, startindex)).ToString();
        }
        else
        {
            UnixTimeString = "無";
        }
        return UnixTimeString;
    }
}
public class EasyCard
{
    public ListViewItem[] ListViewItem;
    public static string[] ReadItem =
        {
            "票卡版號","票卡功能設定","外觀卡號","子區碼","票卡到期日期",
            "交易前票卡錢包餘額","交易前票卡交易序號","卡別","身份別","身份／特種票到期日",
            "區碼","個人身分認證","社福免費搭乘累積優惠點數","社福卡免費搭乘交易日期","Mifare卡號",
            "Mifare卡號長度","舊設備編號","新設備編號","舊服務業者代碼","新服務業者代碼",
            "舊場站代碼","新場站代碼","發卡公司","銀行代碼","忠誠點",
            "轉乘交易序號","轉乘交易日期時間","轉乘交易方式","轉乘交易金額","轉乘交易後餘額",
            "轉乘群組代碼","場站代碼","設備編號","身分證號碼","特種票交易公司代碼",
            "特種票票種","特種票起始日","特種票到期日","特種票期限","特種票已用次數",
            "特種票首次交易日期","段碼","交易時間","上車車站代碼","下車車站代碼",
            "路線代碼","設備編號","里程客運變動區前3個bytes","設備編號","狀態",
            "路線代碼","交易時間","車站代碼","交易金額／預扣金額","搭乘模式",
            "VIP票累積已用點數","搭乘次數旗標","剩餘加值額度","里程計費_累積金額","里程計費_交易時間"
        };
    /***************************************************************************************************/
    public EasyCard()
    {
        ListViewItem = new ListViewItem[ReadItem.Length];
        for (int i = 0; i < ListViewItem.Length; i++)
        {
            ListViewItem[i] = new ListViewItem();
        }
    }
    public ListViewSubItem[] ByteArrayToListViewSubItem(byte[] data)
    {
        ListViewSubItem[] ListViewSubItem = new ListViewSubItem[ListViewItem.Length];
        for (int i = 0; i < ListViewItem.Length; i++)
        {
            ListViewSubItem[i] = new ListViewSubItem();
        }
        ListViewSubItem[0].Text = BitConverter.ToString(data, 5, 1);//票卡版號
        ListViewSubItem[1].Text = BitConverter.ToString(data, 6, 1);//票卡功能設定
        ListViewSubItem[2].Text = BitConverter.ToString(data, 7, 8);//外觀卡號
        ListViewSubItem[3].Text = Convert.ToString(BitConverter.ToUInt16(data, 15));//子區碼
        ListViewSubItem[4].Text = ByteArrayToUnixTimeString(data, 17);//票卡到期日期
        ListViewSubItem[5].Text = Convert.ToString(BitConverter.ToInt16(data, 21));//交易前票卡錢包餘額
        ListViewSubItem[6].Text = Convert.ToString(BitConverter.ToUInt16(data, 24));//交易前票卡交易序號
        ListViewSubItem[7].Text = BitConverter.ToString(data, 27, 1);//卡別
        ListViewSubItem[8].Text = BitConverter.ToString(data, 28, 1);//身份別
        ListViewSubItem[9].Text = ByteArrayToUnixTimeString(data, 29);//身份／特種票到期日
        ListViewSubItem[10].Text = BitConverter.ToString(data, 33, 1);//區碼
        ListViewSubItem[11].Text = BitConverter.ToString(data, 34, 1);//個人身分認證
        ListViewSubItem[12].Text = Convert.ToString(BitConverter.ToUInt16(data, 35));//社福免費搭乘累積優惠點數
        ListViewSubItem[13].Text = ByteArrayToDosDateString(data,37);//社福卡免費搭乘交易日期
        ListViewSubItem[14].Text = BitConverter.ToString(data, 39, 7);//Mifare卡號
        ListViewSubItem[15].Text = BitConverter.ToString(data, 46, 1);//Mifare卡號長度
        ListViewSubItem[16].Text = BitConverter.ToString(data, 47, 4);//舊設備編號
        ListViewSubItem[17].Text = BitConverter.ToString(data, 51, 6);//新設備編號
        ListViewSubItem[18].Text = BitConverter.ToString(data, 57, 1);//舊服務業者代碼
        ListViewSubItem[19].Text = BitConverter.ToString(data, 58, 3);//新服務業者代碼
        ListViewSubItem[20].Text = BitConverter.ToString(data, 61, 1);//舊場站代碼
        ListViewSubItem[21].Text = BitConverter.ToString(data, 62, 2);//新場站代碼
        ListViewSubItem[22].Text = BitConverter.ToString(data, 64, 1);//發卡公司
        ListViewSubItem[23].Text = BitConverter.ToString(data, 65, 1);//銀行代碼
        ListViewSubItem[24].Text = BitConverter.ToString(data, 66, 2);//忠誠點
        ListViewSubItem[25].Text = BitConverter.ToString(data, 101, 3);//轉乘交易序號
        ListViewSubItem[26].Text = ByteArrayToUnixTimeString(data, 104);//轉乘交易日期時間
        ListViewSubItem[27].Text = BitConverter.ToString(data, 108, 1);//轉乘交易方式
        ListViewSubItem[28].Text = Convert.ToString(BitConverter.ToUInt16(data, 109));//轉乘交易金額
        ListViewSubItem[29].Text = Convert.ToString(BitConverter.ToInt16(data, 112));//轉乘交易後餘
        ListViewSubItem[30].Text = BitConverter.ToString(data, 115, 2);//轉乘群組代碼
        ListViewSubItem[31].Text = BitConverter.ToString(data, 117, 2);//場站代碼
        ListViewSubItem[32].Text = BitConverter.ToString(data, 119, 6);//設備編號
        ListViewSubItem[33].Text = BitConverter.ToString(data, 127, 16);//身分證號碼
        ListViewSubItem[34].Text = BitConverter.ToString(data, 143, 3);//特種票交易公司代碼

        ListViewSubItem[35].Text = BitConverter.ToString(data, 146, 1);//特種票票種
        ListViewSubItem[36].Text = ByteArrayToDosDateString(data, 147);//特種票起始日
        ListViewSubItem[37].Text = ByteArrayToDosDateString(data, 149);//特種票到期日
        ListViewSubItem[38].Text = Convert.ToString(data[151]);//特種票期限
        ListViewSubItem[39].Text = Convert.ToString(data[169]);//特種票已用次數

        ListViewSubItem[40].Text = ByteArrayToDosDateString(data, 170);//特種票首次交易日期
        ListViewSubItem[41].Text = BitConverter.ToString(data, 172, 1);//段碼
        ListViewSubItem[42].Text = ByteArrayToUnixTimeString(data, 173);//交易時間
        ListViewSubItem[43].Text = BitConverter.ToString(data, 177, 1);//上車車站代碼
        ListViewSubItem[44].Text = BitConverter.ToString(data, 178, 1);//下車車站代碼

        ListViewSubItem[45].Text = BitConverter.ToString(data, 179, 2);//路線代碼
        ListViewSubItem[46].Text = BitConverter.ToString(data, 181, 6);//設備編號
        ListViewSubItem[47].Text = BitConverter.ToString(data, 221, 3);//里程客運變動區前3個bytes
        ListViewSubItem[48].Text = BitConverter.ToString(data, 224, 6);//設備編號
        ListViewSubItem[49].Text = BitConverter.ToString(data, 230, 1);//狀態

        ListViewSubItem[50].Text = BitConverter.ToString(data, 231, 2);//路線代碼
        ListViewSubItem[51].Text = ByteArrayToUnixTimeString(data, 233);//交易時間
        ListViewSubItem[52].Text = BitConverter.ToString(data, 237, 2);//車站代碼
        ListViewSubItem[53].Text = Convert.ToString(BitConverter.ToUInt16(data, 239));//交易金額／預扣金額
        ListViewSubItem[54].Text = BitConverter.ToString(data, 242, 1);//搭乘模式

        ListViewSubItem[55].Text = Convert.ToString(BitConverter.ToUInt16(data, 243));//VIP票累積已用點數
        ListViewSubItem[56].Text = BitConverter.ToString(data, 245, 1);//搭乘次數旗標
        ListViewSubItem[57].Text = Convert.ToString(BitConverter.ToUInt16(data, 247));//剩餘加值額度
        ListViewSubItem[58].Text = Convert.ToString(data[250]);//里程計費_累積金額／預扣金額
        ListViewSubItem[59].Text = ByteArrayToDosDateString(data, 251);//里程計費_交易時間
        return ListViewSubItem;
    }
    private string ByteArrayToUnixTimeString(byte[] data, int startindex)
    {
        string UnixTimeString;
        if (BitConverter.ToUInt32(data, startindex) > 0)
        {
            UnixTimeString = new DateTime(1970, 1, 1, 0, 0, 0).AddSeconds(BitConverter.ToInt32(data, startindex)).ToString();
        }
        else
        {
            UnixTimeString = "無";
        }
        return UnixTimeString;
    }
    private string ByteArrayToDosDateString(byte[] data, int startindex)
    {
        byte[] DosDateArray = new byte[2];
        ushort DosDateUShort;
        string DosDateString;
        Array.Copy(data, startindex, DosDateArray, 0, 2);
        Array.Reverse(DosDateArray);
        DosDateUShort = BitConverter.ToUInt16(DosDateArray, 0);
        if (DosDateUShort > 0)
        {
            DosDateString = new DateTime(1980, (DosDateUShort & 0x01E0) >> 5, DosDateUShort & 0x001F).AddYears((DosDateUShort >> 9)).ToString("yyyy/MM/dd");
        }
        else
        {
            DosDateString = "無";
        }
        return DosDateString;
    }
}