using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Net;
#region --- лицензия GNU
//                           LICENSE INFORMATION
//********************************************************************************************
// Sapphire 400 Automated Clinical Analyzer Driver, 
// serial port & SQL communications. Version 1.05.211.
// This programme is for managing data from serial port & its transmission to MS-SQL Server
// according to the rules described in "Automated Clinical Analyzer. BiOLiS 24i Premium.
// Bi-directional Communicatiom Specifications. Version 1.07" 
// by Tokyo Boeki Medical System Ltd.
//
// Copyright (C) 2020 Vladimir A. Maltapar
// Email: maltapar@gmail.com
// Created: 04 April 2020
//
// This program is free software: you can redistribute it and/or modify  it under the terms of
// the GNU General Public License as published by the Free Software Foundation, 
// either version 3 of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License along with this program.
// If not, see <http://www.gnu.org/licenses/>.
//********************************************************************************************
#endregion --- лицензия GNU

namespace Sapphire
{
    public partial class FormSap : Form
    {
        #region --- Общие параметры
        // параметры в строчках ini-файла 
        private static string connStr;      // строка коннекта к SQL
        private static int    Analyzer_Id = 911;      // для Sapphire400 =11, он берётся из ini-файла.
        private static string Analyzer_Name = "Testname"; // SIEMENS или SAPPHIRE
        private static string sModes;       // строка списка режимов работы 
        private static string PathLog;      // путь к лог-файлу с вычисленными каталогом и датой ...\GGGG-MM\BeckLog2019-08-06.txt 
        private static string PathLogDir;   // путь к лог-файлу, заданный в параметрах ini-файла  
        private static string PathErrLog;   // путь к лог-файлу ошибок 
        private string PathIni;             // Путь к ini-файлу (там, откуда запуск)
        private static int LogInterval = 30;// интервал логирования работы в ожидании, в сек.
        private string inputString = "";    // полученные данные - для парсинга
        // глобальные
        private static DateTime dt0 = DateTime.Now; // время старта 
        private static DateTime dtm = dt0;          // для измерения времени выполнения запроса 
        private readonly string sTimeStart = dt0.ToString("dd-MM-yyyy HH:mm:ss");
        private int    NumberOfStart = 0;
        private readonly string UserName = System.Environment.UserName;
        private readonly string ComputerName = System.Environment.MachineName;
        //private string myIP="" ;
        private string AppVer; // версия из AutoVersion2
        private static string AppName;      // static т.к. исп. в background-процессе
        private static string qkreq = "";   // запрос о работе ("кукареку")
        private static string sHeader1 = "";// строка заголовка - назначение, для чего.
        private string msg, msg0, msg1, msg2;
        private string ComPortName;
        private string ENQ = "\x05", ACK = "\x06", NAK = "\x15", ETB = "\x17" 
                      ,STX = "\x02", ETX = "\x03", EOT = "\x04", CR  = "\n", LF = "\r"; // управляющие коды в протоколе обмена для СОМ-порта
        private Random rnd = new Random(3213);
        private int partNo = 0;   // i-я часть принятых данных
        private string KodJournal = "0";  // 1 - Экстренный, 2 - Дежурный, 3 - Плановый
        #endregion --- Общие параметры
        #region --- для SQL
        // для результата Select
        private static string dateDone = "2019-12-31 23:59";  // дата-время выполнения анализа по часам на анализаторе
        private static string dateDone999 = "31-12-2019 23:59:00.000";  // дата-время выполнения анализа для SQL
        private string st1_0 = "Insert into AnalyzerResults (Analyzer_id, HostName, ResultDate, CntParam";
        private string st1_1 = "";
        private string st2_0 = ""; //$") values ({Analyzer_Id}, host_name(), GetDate(), {CntParam}, {ResultText}";    // для строки Insert...
        private string st2_1 = "";                          // для строки Insert...
        private string st0 = "";
        private static string HistoryNumber = "0", sResult = "", sEdizm = "", sRefVal = "", sFlag = "", dt_Done = "", ResultDate="", ResultText="";
        private static int    CntParam;
        private static int    MaxCntParam;       // = 22 на 2020-04-24
        private static int    MaxSapphireResult; // = 12 на 2020-04-24
        private static bool   IsControl = false; // признак - текущая принятая запись - это контроли.
        private string        patternHistN = ""; // для выделения HistoryNumber  @"\d{0,6}";  - от 0 до 6 цифр максимум! 
        private int kRecord = 0; // кол-во полученных записей
        #endregion --- для SQL
        public FormSap()
        {
            InitializeComponent();
            ReadParmsIni();         // берём параметры из ini-файла 
            //MessageBox.Show("040", "Check point number:", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            FormIni();              // установки на форме, которые не делает Visual Studio - IP-адреса          
            PortIni();
        }
        #region --- методы на форме: timer, ...
        private void FormIni()          // мой начальный вывод на форму, что прочитали из ini-файла
        {
            var host = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
            var list_adr = from ip in host where ip.ToString().Length < 16 select ip.ToString();
            string IP = "";
            foreach (string st in list_adr) IP += st + " ";
            Lbl_IP.Text = "IP: " + IP;

            CmbTest.SelectedIndex = 0;   // первый (нулевой) элемент - текущий, видимый.

            if (File.Exists($"{PathIni}\\Pic.png"))
            {
                this.Pic1.Image = new Bitmap($"{PathIni}\\Pic.png");  // на форме картинка - для различных приложений должна быть другая!
            }

            notifyIcon1.Visible = false; // невидимая иконка в трее
            // добавляем событие по 2-му клику мышки, вызывая функцию  NotifyIcon1_MouseDoubleClick
            this.notifyIcon1.MouseDoubleClick += new MouseEventHandler(NotifyIcon1_MouseDoubleClick);

            // добавляем событие на изменение окна
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.ShowInTaskbar = false;

            // запуск свёрнутым в трей
            this.WindowState = FormWindowState.Minimized;
            this.Hide();

            timer1.Interval = LogInterval * 1000;
            timer1.Start();
        }
        private void NotifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            notifyIcon1.Visible = false;
            WindowState = FormWindowState.Normal;
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon1.Visible = true; // иконка в трее видимая
                notifyIcon1.ShowBalloonTip(200); // 2 секунды показать предупреждение о сворачивании в трей
            }
            else if (FormWindowState.Normal == this.WindowState)
            { notifyIcon1.Visible = false; }
        }
        private void Timer1_Tick(object sender, EventArgs e)
        {
            // текущую дату-время записать в файл %PathErrLog%\LogTimer.txt (для мониторинга работы прораммы)
            dt0 = DateTime.Now;
            string fnLogAlive = PathErrLog + @"\LogAlive.txt";
            fnLogAlive = Path.GetFullPath(fnLogAlive);
            FileStream fn = new FileStream(fnLogAlive, FileMode.Create);  // FileMode.Append
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"{ss} ID: {Analyzer_Id}");
            sw.Close();
        }
        #endregion методы на форме: timer, ...
        #region --- методы для Com-порта: инициализация (PortIni) и чтения (Sp_DataReceived и Si_DataReceived)
        public SerialPort _serialPort;
        // Делегат используется для записи в UI control из потока не-UI
        private delegate void SetTextDeleg(string text);
        // Все опции для последовательного устройства могут быть отправлены через конструктор класса SerialPort
        // PortName = "COM1", Baud Rate = 19200, Parity = None, Data Bits = 8, Stop Bits = One, Handshake = None
        //public SerialPort _serialPort = new SerialPort("COM2", 9600, Parity.None, 8, StopBits.One);

        private void PortIni() // инициализация: проверка порта, ...
        {
            if (sModes.IndexOf("NoComPort") >= 0)
            {
                Stat1.Text = $"Тест - без Com-порта.";
                Stat1.ForeColor = Color.DarkRed;
                return;
            }
            _serialPort = new SerialPort(ComPortName, 9600, Parity.None, 8, StopBits.One)
            {
                // ReadTimeout = 500; // в милисекундах
                Handshake = Handshake.None, // возможные значения: None OnXOff RequestToSendXOnXOff RequestToSend
                ReadTimeout  = 500,
                //Encoding= Encoding.ASCII,   // 2020-04-20
                WriteTimeout = 500
            };
            //_serialPort.Encoding = Encoding.ASCII;
            _serialPort.Encoding = Encoding.GetEncoding(1251);  // 2020-04-21
            Add_RTB(RTBout, $"\nCodePage: {_serialPort.Encoding.CodePage}.\n", Color.Blue);
            //_serialPort.Encoding.CodePage.ToString()
            _serialPort.DataReceived += new SerialDataReceivedEventHandler(Sp_DataReceived);
            // Открытие последовательного порта
            // Перед попыткой записи убедимся, что порт открыт.
            try
            {
                if (!(_serialPort.IsOpen))
                {
                    _serialPort.Open();
                    Stat1.Text = $"Открыт порт {_serialPort.PortName}.";
                }
                else
                    Stat1.Text = $"Уже открыт {_serialPort.PortName}!";
            }
            catch (Exception ex)
            {
                string mes = "Ошибка открытия порта";
                WErrLog(mes);
                ExitApp($"{mes} {_serialPort.PortName}.\n{ex.Message}", 2);
            }
        }

        void Sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string dataCom = _serialPort.ReadExisting();
            // Привлечение делегата на потоке UI и отправка данных, которые были приняты привлеченным методом.
            // Метод "Si_DataReceived" будет выполнен в потоке UI, который позволит заполнить текстовое поле TextBox.
            this.BeginInvoke(new SetTextDeleg(Si_DataReceived),
                             new object[] { dataCom });
        }

        private void Si_DataReceived(string receivedData) // обработка полученных данных
        {
            int kb = receivedData.Count();
            partNo++; // номер очередной части принятых данных
            inputString += receivedData;

            //byte[] receivedByte = Encoding.Default.GetBytes(receivedData);
            //qkreq = "test_receivedByte";

            Stat1.Text = " Приём данных...";
            Stat1.ForeColor = Color.Red;
            //Pic_Status.Visible = true;  //ToDo сделать изменяемую картинку в статусной строке
            dtm = DateTime.Now;

            if (sModes.IndexOf("Log_receivedParts") != -1)
                WLog($"receivedData: {receivedData}");

            msg = $"{partNo}-я часть приёма, байт: {kb}.";
            if (sModes.IndexOf("Show_receivedParts") != -1)
            {
                Add_RTB(RTBout, $"{dtm} {msg}", Color.DarkBlue);
                Add_RTB(RTBout, $"\n{receivedData}", Color.Black);
            }
            if (receivedData.IndexOf(ENQ) != -1) // часть данных c ENQ
            {
                _serialPort.Write(ACK);   // подтверждение <ACK>
                Add_RTB(RTBout, $"\n{dtm}, ответ на данные с <ENQ>: <ACK>.\n", Color.DarkGreen);
            }
            int iSTX = receivedData.IndexOf(STX);
            if ((iSTX != -1) & (kb >= 3))  // есть <STX>
            {
                _serialPort.Write(ACK);   // подтверждение <ACK>
                Add_RTB(RTBout, $"{dtm} есть <STX>, ответ: <ACK>.\n", Color.DarkGreen);
            }
            int iEOT = receivedData.IndexOf(EOT);
            if (iEOT == kb-1) //  есть признак конца <EOT> 04h и это последний байт   ## (iEOT != -1)
            {
                Add_RTB(RTBout, $"{dtm} есть <EOT>\n", Color.DarkGreen);
                if (iEOT > kb - 1)    // есть данные за концом    ## (kb != 1) должен приходить 1 байт
                {
                    msg = "ERR: получен признак <EOT> и ещё данные, они не обработаны! :(";
                    Add_RTB(RTBout, $"{dtm} {msg}\n", Color.Red);
                    //MessageBox.Show(msg, "Внимание! Ошибка в двнных!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    msg += $"\n partNo={partNo}\n iEOT={iEOT}\n После <EOT>:{receivedData.Substring(iEOT)}\n receivedData={receivedData}";
                    WErrLog(msg);
                    WLog(msg);
                    CntParam = 0;  // чтобы не писать в SQL
                }
                else
                {
                    ParseResults(inputString);  // вызвать обработчик ...
                }
                // ToSQL...
                if (CntParam == 0) // пропускаем при количестве анализов=0
                {
                    msg = "Кол-во анализов = 0: пропускаем, в SQL не пишем.";
                    Add_RTB(RTBout, $"{dtm} {msg}\n", Color.DarkGreen);
                    WLog(msg);
                }
                else if (IsControl) // контроли пропускаем (пока: 2020-04-16)
                {
                    msg = "Это контроли: их пока пропускаем, но надо писать отдельно!";
                    Add_RTB(RTBout, $"{dtm} {msg}\n", Color.DarkGreen);
                    WLog(msg);
                }
                else if (CntParam > MaxCntParam) // возможно, это повторы (2020-04-16) 
                {
                    // записать собщение для лаборанта...
                    msg = $"Больше {MaxCntParam} анализов! Это повторы: их пока пропускаем (всего {MaxCntParam} поля для приёма в LabAutoResult!)";
                    Add_RTB(RTBout, $"{dtm} {msg}\n", Color.DarkGreen);
                    WLog(msg);
                }
                else //  пишем в MS-SQL
                {
                    msg = "Результаты: пишем в SQL.";
                    Add_RTB(RTBout, $"{dtm} {msg}\n", Color.DarkGreen);
                    WLog(msg + " " + st0);
                    ToSQL(st0);
                }

                if (HistoryNumber == "0" && CntParam > 0) // и не запрос Query ALL
                {
                    // записать собщение для лаборанта...
                }

                HistoryNumber = "0"; sResult = ""; sEdizm = ""; sRefVal = ""; sFlag = ""; dt_Done = "";
                CntParam = 0; ResultDate = "";

                inputString = ""; partNo = 0;   // очистить inputString  для повторного приёма данных
            }
            else
            {
                // ======== ПРОЧЕЕ - склеивание частей полученных данных =========
                Add_RTB(RTBout, $"{dtm} ... ещё пришло ...\n", Color.DarkGreen);
            }
            Stat1.Text = "... ждём-с...";
            Stat1.ForeColor = Color.Green;
        }
        #endregion --- методы для Com-порта: инициализация (PortIni) и чтения (Sp_DataReceived и Si_DataReceived)
        // ---
        #region --- Парсинг данных из строки и формирование строки SQL Insert
        private void ParseResults(string input) // ================= ПАРСИНГ ПОЛУЧЕННЫХ ДАННЫХ ===================
        {
            Stat1.Text = "парсинг";
            Stat1.ForeColor = Color.DarkBlue;
            dtm = DateTime.Now;
            Add_RTB(RTBout, $"{dtm} Начало парсинга. Получено байт: {inputString.Count()}.\n", Color.DarkViolet);
            Add_RTB(RTBout, $"{inputString}", Color.Black);
            IsControl = false;
            st1_1 = ""; st2_1 = "";
            CntParam = 0;
            inputString = Regex.Replace(inputString, ETX + "..", String.Empty);  // удаление ETX и следующих за ним двух байт контрольной суммы - они не нужны мне.

            inputString = Regex.Replace(inputString, ETB + ".." + CR + LF + STX + ".", String.Empty);
            /* строка выше - это замена управляющих символов, если данные разбиты на несколько фреймов: 
                 <STX> [FN] [TEXT] <ETB> [C1][C2] <CR> <LF> Intermediate frame
                 <STX> [FN] [TEXT] <ETX> [C1][C2] <CR> <LF> Termination frame
              Вот пример:
                [STX]3O|1|01040218001|^1^1|^^^1^T-BIL^0・・・¥^^^18^CPK^0¥^^^19[ETB]??[CR][LF]
                [STX]4^AMY^0|R||||||N||||Serum||||||||||F[CR][ETX]??[CR][LF]
            */
            inputString = Regex.Replace(inputString, CR + LF, String.Empty);     // удаление CR+LF
            Regex rg = new Regex(patternHistN);   //string patternHistN = @"\d{0,6}";   // для выделения HistoryNumber - в параметрах ini-файла! 
            MatchCollection matched;
            string sHist;  // " 25781/ ТЕСТОВЫЙ Ая АРО1";   "?16512 ^??????? ?.?. ???1"  - это первоначальная строка с номером ИБ, ФИО, ...
            string[] sRecord = inputString.Split(new string[] { CR, LF, ENQ, STX, EOT }, StringSplitOptions.RemoveEmptyEntries);
            int ks = sRecord.Count();

            for (int i = 0; i < ks; i++)  // по количеству строк в переданных результатах пациента
            {
                string[] sField = sRecord[i].Split('|');
                int kField = sField.Count();
                // CheckSubLength
                if (sRecord[i].Length<2)
                {
                    msg1 = $"ERR: Ошибка в данных - неизвестный тип записи в {i}-й строке:{sRecord[i]}.";
                    Add_RTB(RTBout, $"\n{dtm} {msg1}.\n", Color.Red);
                    WLog(msg1);
                    WErrLog(msg1);
                    continue;
                }
                string sRecordType = sRecord[i].Substring(1, 1);  // CheckSubLength

                if (sRecordType == "P" & Analyzer_Name.ToUpper() != "SIEMENS")
                {   // Patient  "2P|1|2020040150101|||19576 ^??????? ?.?   ???|||U|||||"
                    //           0  1 2            345                        67 
                    if (kField == 2)
                    {
                        IsControl = true;
                        Add_RTB(RTBout, $"\n - Это контроли, их пропускаем...", Color.DarkGreen);
                        break;  // это  контроли, их пока не обрабатываем (но надо, я их в Науке собирал для внутрилабораторных отчётов)
                    }
                    sHist = sField[5].Trim();
                    matched = rg.Matches(sHist);
                    //HistoryNumber = "0";
                    //for (int j = 0; j < matched.Count; j++)
                    //{
                    //    if (matched[j].Value.Length != 0)
                    //    {
                    //        HistoryNumber = matched[j].Value;
                    //        break;  // результат д.б. в первой непустой
                    //    }
                    //}
                    if (matched.Count>0)
                        HistoryNumber = matched[0].Value;
                    if (HistoryNumber.Length==0)
                        HistoryNumber = "0";
                }

                if (sRecordType == "O" & Analyzer_Name.ToUpper() == "SIEMENS")  // выделяем номер истории - только для SIEMENS
                {   // Order "6O|2|2-23636||^^^1|R||||||||||||||||||||X"
                    //        0  1 2       
                    sHist = sField[2].Trim();
                    int im = sHist.IndexOf("-");
                    if (im == -1 | im == sHist.Length - 1)
                    {
                        //Add_RTB(RTBout, $"\nНет номера истории!\n", Color.Red);
                        HistoryNumber = "0";
                        // return; BugFix 2021-06-03 пишем в SQL даже если нет номера истории.
                    }
                    else
                    {
                        HistoryNumber = sHist.Substring(im + 1); // после первого минуса и до конца строки должен быть номер истории. 2021-05-31.
                    }
                }

                if (sRecordType == "O" & Analyzer_Name == "SAPPHIRE")  // Serum - сыворотка, Urine, СМЖ
                { // 3O|1||^1^77|^^Ю1^??ї?ї?ї^0\^^Ю5^??.??ї^0\^^Ю7^??ї?ї-?^0\^^Ю8^??ї?ї?ї?^0\^^Ю9^??ї^0\^^Ю10^??ї^0\^^Ю11^??ї?ї?ї^0\^^Ю16^??ї.??^0|R||ь|ь|ь|ь|Serum||ь|ь|ь|ь|F
                  //  0 1 2 3    4                                                                                                                 5 67 8 9 10 11  12   
                        
                }

                if (sRecordType == "R")
                {   // Result "4R|1|^^^1^???????^0| 4.7804|?????/?| 4.2000 TO  6.4000|N||F||||20200415165509" для Sapphire
                    //         0  1 2             3        4       5                  6 78 9012 
                    // следующие результаты - для SIEMENS:
                    // Result "4R|1|^^^4^PT.sec.TS^^^F|14.400|sec||H||F||||20210531095944"
                    //          0 1 2                  3      4    6  8    12 
                    // Result "7R|1|^^^1^aPTT.PSL^^^F|41.108|sec||H||F||||20210531095728"
                    //          0 1 2                 3      4    6  8    12 
                    // Result "0R|1|^^^2^PT.INR.Cal.TS.MC^^^F|1.153|||H||F||||20210531095944"
                    //          0 1 2                         3    4  6  8    12 
                    // Result "2R|1|^^^5^BC.TT^^^F|16.561|sec||||F||||20210531095728"
                    //          0 1 2              3      4    6 8    12 
                    // Result "5R|1|^^^6^Fib^^^F|4.675|g/L||H||F||||20210531101225"
                    //          0 1 2            3     4    6  8    12 

                    string sKod = sField[2];  // "^^^1^???????^0"  или пустая строка!!!
                    Regex reg_kod = new Regex(@"\d+");  // найти первые n цифр: Regex(@"\d+"); первые две цифры: Regex(@"\d{1,2}");
                    MatchCollection matched_kod = reg_kod.Matches(sKod); // fix 2021-02-10
                    if (matched_kod.Count > 0)
                        sKod = matched_kod[0].Value.ToString();
                    else
                    {
                        msg1 = $"ERR: Нет кода анализа! kField={kField}, sRecord({i})={sRecord[i]}";
                        Add_RTB(RTBout, $"\n{dtm} {msg1}.\n", Color.Red);
                        WLog(msg1);
                        WErrLog(msg1);
                    }

                    sResult = sField[3].Trim();
                    sResult = sResult.Replace("?", "");
                    sResult = sResult.Replace("/", "");
                    sEdizm = sField[4].Trim();
                    sRefVal = sField[5].Trim();
                    sRefVal = Regex.Replace(sRefVal, @"\s+", " ");       // заменяет несколько подряд идущих пробелов одинарными
                    if (sRefVal.Length > 16)
                        sRefVal = "длина больше 16";
                    
                    if (Analyzer_Name == "SIEMENS")
                        sFlag = sField[7];
                    else
                        sFlag = sField[6];

                    if (kField >= 13)   // для контроля индекса dt_Done = sField[12];
                        dt_Done = sField[12];
                    else if (kField == 12)
                        dt_Done = sField[11];
                    else if (kField == 11)
                        dt_Done = sField[10];
                    else //  (kField < 11)                   
                    {
                        msg1 = $"ERR: Мало полей! (индекс!) kField={kField}, sRecord({i})={sRecord[i]}";
                        Add_RTB(RTBout, $"\n{dtm} {msg1}.\n", Color.Red);
                        WLog(msg1);
                        WErrLog(msg1);
                        dt_Done = "";
                    }

                    CntParam++;
                    if (CntParam > MaxCntParam)
                    {
                        msg1 = $"ERR: CntParam ({CntParam}) > ({MaxCntParam}) MaxCntParam! Хвост с результатами не записан - некуда!";
                        Add_RTB(RTBout, $"\n{dtm} {msg1}.\n", Color.Red);
                        WLog(msg1);
                        WErrLog(msg1);
                        break;  // Что ещё нужо?  
                    }
                    if (CntParam > MaxSapphireResult)
                    {
                        msg1 = $"ERR: Количество анализов больше максимального ({MaxSapphireResult})!";
                        Add_RTB(RTBout, $"\n{dtm} {msg1}.\n", Color.Red);
                        WLog(msg1);
                    }
                    st1_1 += $", ParamName{CntParam}, ParamValue{CntParam}, ParamRef{CntParam}";    // + ParamMgr{CntParam}
                    st2_1 += $", '{sKod}', '{sResult}', '{sRefVal}'";
                    Add_RTB(RTBout, $"\nКод:{sKod}, результат:{sResult}, реф.знач:{sRefVal}, выполнен:{dt_Done}.", Color.DeepPink);
                }
            }

            // формирование строки SQL Insert...
            ResultText = inputString;
            st2_0 = $", HistoryNumber, ResultText) values ({Analyzer_Id}, host_name(), GetDate(), {CntParam} ";    // для строки Insert...
            st0 = st1_0 + st1_1 + st2_0 + st2_1 + $", '{HistoryNumber}','{ResultText}' );";
            dtm = DateTime.Now;
            Add_RTB(RTBout, $"\n{dtm} Конец парсинга [{kRecord++}]. HistoryNumber:{HistoryNumber}.\n", Color.DarkViolet);
            WLog($"Получено для парсинга {inputString.Count()} байт:\n{inputString}");
            Stat1.Text = "конец парсинга";
           }
        #endregion --- Парсинг данных из строки и формирование строки SQL Insert        
        private void ToSQL(string st) // запись в MS-SQL сформированной строки
        {
            using (SqlConnection sqlConn = new SqlConnection(connStr))
            {
                //SqlCommand sqlCmd = new SqlCommand("INSERT INTO [AnalyzerResults]([Analyzer_Id],[ResultText],[ResultDate],[Hostname],[HistoryNumber])VALUES(@AnalyzerId,@ResultText,GETDATE(),@PCname,@HistoryNumber)", sqlConn);
                SqlCommand sqlCmd = new SqlCommand(st, sqlConn);
                try
                {
                    Stat1.Text = "запись в SQL...";
                    Stat1.ForeColor = Color.Blue;
                    sqlCmd.CommandType = System.Data.CommandType.Text;
                    sqlConn.Open();
                    sqlCmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    string mes = "Ошибка при записи в SQL.";
                    WErrLog(mes + "\n" + ex.ToString()); // в файл ошибок...
                    WLog   (mes + "\n" + ex.ToString()); // в лог файл тоже!
                    ExitApp(mes, 3);
                }
            }
        }
        #region --- CheckExist Проверка на уже существующий такой же сегодняшний результат
        private Int32 CheckExist(string ResText) // функция проверки существующего в SQL
        {
            Int32 icnt = -1;
            string dtc = DateTime.Today.ToString("yyyyMMdd");
            DateTime dtek = DateTime.ParseExact(dtc, "yyyyMdd", null);
            string SQL_Select = "select[dbo].[CheckExistBeckmanData](@s, @dt) as cnt";
            using (SqlConnection sqlConn = new SqlConnection(connStr))
            {
                sqlConn.Open();
                SqlCommand sqlCmd = new SqlCommand(SQL_Select, sqlConn);
                sqlCmd.CommandType = System.Data.CommandType.Text;
                sqlCmd.Parameters.AddWithValue("@dt", dtek);
                sqlCmd.Parameters.AddWithValue("@s", ResText);
                Object o = sqlCmd.ExecuteScalar();
                if (o == DBNull.Value)
                    icnt = 0;
                else
                    icnt = Convert.ToInt32(o);
                return (icnt);
            }
        }
        #endregion --- CheckExist Проверка на уже существующий такой же сегодняшний результат
        private void ReadParmsIni()     // читать настройки из ini-файла 
        {
            //FIXME for WinXP1
            //MessageBox.Show("030", "Check point number:", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            PathIni = Application.StartupPath;
            AppName = AppDomain.CurrentDomain.FriendlyName;
            AppName = AppName.Substring(0, AppName.IndexOf(".exe"));
            string pathIniFile = PathIni + @"\" + $"\\{AppName}" + ".ini";
            if (!File.Exists(pathIniFile))
            {
                string errmsg = "Не найден файл " + pathIniFile + "\n Работа завершается!";
                MessageBox.Show(errmsg, " Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Environment.Exit(1);
            }

            INIManager iniFile = new INIManager(pathIniFile);   // cоздание объекта для работы с ini-файлом

            // секция [Description]
            sHeader1 = iniFile.GetPrivateString("Description", "Header1"); // получить значение  из секции Description по ключу Header1
            RtbHeader1.Text = sHeader1;
            this.Text = "  " + sHeader1;    // заголовок в основном окне  

            // секция [Connection]
            string str = iniFile.GetPrivateString("Connection", "Analyzer_Id").Trim();
            Analyzer_Id = Convert.ToInt32(str);
            Analyzer_Name = iniFile.GetPrivateString("Connection", "Analyzer_Name").Trim();
            connStr = iniFile.GetPrivateString("Connection", "DbSQL").Trim(); // получить значение  из секции 1.. по ключу 2..
            string nameSQLsrev = connStr.Substring(0, connStr.IndexOf(';'));
            ComPortName = iniFile.GetPrivateString("Connection", "ComPortNo").Trim();
            // секция [Check]
            /*
            [Check]
            patternHistN=\d{0,6}
            MaxCntParam=22
            MaxSapphireResult=12
            LogInterval=20
            */
            patternHistN = iniFile.GetPrivateString("Check", "patternHistN").Trim(); // шаблон для выделения номера истории по RegEx patternHistN=\d{0,6}
            str = iniFile.GetPrivateString("Check", "MaxCntParam").Trim();
            MaxCntParam = Convert.ToInt32(str);
            str = iniFile.GetPrivateString("Check", "MaxSapphireResult").Trim();
            MaxSapphireResult = Convert.ToInt32(str);
            str = iniFile.GetPrivateString("Check", "LogInterval").Trim();
            LogInterval = Convert.ToInt32(str);
            // секция [LogFiles]
            PathLogDir = iniFile.GetPrivateString("LogFiles", "PathLogDir");
            //SetPathLog();  // Fixed 2021-02-11
            PathErrLog = iniFile.GetPrivateString("LogFiles", "PathErrLog");

            // секция [Modes]  Режимы работы 
            sModes = iniFile.GetPrivateString("Modes", "mode1");
            /* Пример содержимого:
            mode1 = (Log_Receive),(...), ...
            */

            // секция [Comments]
            string sNumberOfStart = iniFile.GetPrivateString("Comments", "StartNo");
            NumberOfStart = Convert.ToInt32(sNumberOfStart);
            NumberOfStart++;
            sNumberOfStart = Convert.ToString(NumberOfStart);
            iniFile.WritePrivateString("Comments", "StartNo", sNumberOfStart);
            iniFile.WritePrivateString("Comments", "LastStartTime", dtm.ToString());// записать значение в секции Comments по ключу xx
            
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string fileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).FileVersion;
            AppVer = "Версия " + version + " от " + fileVersion;
            WLog($"--- Запуск {AppName} {ComputerName} {UserName} {AppVer}");
        }
        // ---
        #region --- ( Easter eggs :))
        private void Pic1_Click(object sender, EventArgs e)
        {
            Pic1.Visible = !Pic1.Visible;
        }
        #endregion --- ( Easter eggs :))
        // ---
        #region ---Меню + методы Wlog, WErrLog, WTest;  SetPathLog, Add_RTB, ExitApp, ...
        private static void SetPathLog()     // нужна, если программа работает много дней и меняется текущая дата //2019-08-06
        {
            dtm = DateTime.Now;
            PathLogDir = Path.GetFullPath(PathLogDir + @"\.");
            string PathLogGodMes = PathLogDir + @"\" + dtm.ToString("yyyy-MM");
            if (!Directory.Exists(PathLogGodMes))
                Directory.CreateDirectory(PathLogGodMes);
            PathLog = PathLogGodMes + @"\" + $"{AppName}_" + dtm.ToString("yyyy-MM-dd") + ".txt";
        }
        private static void WLog(string st) // записать в лог FLog
        { 
            dtm = DateTime.Now;
            PathLogDir = Path.GetFullPath(PathLogDir + @"\.");
            string PathLogGodMes = PathLogDir + @"\" + dtm.ToString("yyyy-MM");
            if (!Directory.Exists(PathLogGodMes))
                Directory.CreateDirectory(PathLogGodMes);
            PathLog = PathLogGodMes + @"\" + $"{AppName}_" + dtm.ToString("yyyy-MM-dd") + ".txt";
            FileStream fn = new FileStream(PathLog, FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"{ss} {st}");
            sw.Close();
        }
        private static void WErrLog(string st) // записать ErrLog
        {
            string fnPathErrLog = PathErrLog + @"\Log_ERR.txt";
            fnPathErrLog = Path.GetFullPath(fnPathErrLog);
            FileStream fn = new FileStream(fnPathErrLog, FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"\n{ss} {st}");
            sw.Close();
        }
        private static void Add_RTB(RichTextBox rtbOut, string addText)
        {
            Add_RTB(rtbOut, addText, Color.Black);
        }
        private static void Add_RTB(RichTextBox rtbOut, string addText, Color myColor)
        {
            Int32 p1, p2;
            p1 = rtbOut.TextLength;
            p2 = addText.Length;
            rtbOut.AppendText(addText);
            rtbOut.Select(p1, p2);
            rtbOut.SelectionColor = myColor;
            // 1 rtbOut.Select(0, 0);
            // 2 rtbOut.Select(p1 + p2, 0);
            // 2 rtbOut.AppendText("");
            rtbOut.SelectionStart = rtbOut.Text.Length;
            rtbOut.ScrollToCaret();
            // или: rtbOut.Select(p1, p2);
            //      SendKeys.Send("^{END}");  // это прокрутка в конец :)
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e) // Закрытие формы - FormClosing
        {
            msg = "Завершение работы.";
            DialogResult result = MessageBox.Show("Вы действительно хотите завершить работу? \n\n   Результаты передаваться не будут!"
                , msg, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                WLog("--- передумал выходить :) " + e.CloseReason.ToString()
                    + " " + result.ToString()); // UserClosing No
                e.Cancel = true;
                return;
            }
            WLog("--- " + msg);
            Environment.Exit(0);
        }
        private static void ExitApp(string mess, int ErrCode = 1001) // Завершение работы по ошибке.
        {
            string title = "Аварийное завершение работы.";
            WErrLog($"{title}\n{mess}");
            MessageBox.Show(mess, title
                , MessageBoxButtons.OK, MessageBoxIcon.Stop);
            Environment.Exit(ErrCode);
        }
        private static void ExitApp(string mess) // Нормальное завершение работы (по кнопке Х)
        {
            WLog("--- " + mess);
            Environment.Exit(0);
        }
        private void НастройкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Меню/Сервис/Настройки
            // ...
        }
        // --- // Меню / Сервис / Параметры в.ini - файле:
        private void ПараметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Меню / Сервис / Параметры в.ini - файле:
            string s = "";
            string nameSQLsrv = connStr.Substring(0, connStr.IndexOf(';'));
            string sep = $"---------------------------------------------------------------------------\n";
            // 2020-04-07 из надстойки Automatic Version 2
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var fileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).FileVersion;
            var productVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).ProductVersion;
            s += $"     {sHeader1}\n";
            s += $"{AppName}.  {AppVer}\n";
            s += sep;
            s += $"Время старта: {sTimeStart}\n";
            s += $"ComputerName: {ComputerName}, UserName: {UserName}\n";
            //s += $"IP: {myIP}\n";
            s += sep;
            s += $"путь к логам:\n";
            s += $"PathLog: {PathLog}\n";       // текущий (с датой)
            s += $"PathLogDir: {PathLogDir}\n"; // начальный (в .ini)
            s += $"PathErrLog: {PathErrLog}\n";
            s += $"PathIni: {PathIni}\n";
            s += sep;
            s += $"Режимы работы: {sModes}\n";
            s += $"SQL: {nameSQLsrv}, Analyzer_Id: {Analyzer_Id}, Analyzer_Name: {Analyzer_Name}.\n";
            s += $"Com-порт: {ComPortName}.\n";
            s += $"Интервал логирования: {LogInterval} сек.\n";
            s += sep+"\n\n\n\n"; 
            DialogResult result = MessageBox.Show(s, "  Параметры в .ini-файле:"
                , MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion --- методы Wlog, WErrLog;  SetPathLog, Add_RTB, ExitApp...
        #region --- Действия по кнопкам на форме
        private void Btn_Clear_Click(object sender, EventArgs e) // Кнопка: очистить форму
        {
            RTBout.Clear();
            partNo = 0;
            kRecord = 0;
        }
        private void Btn_ACK_Click(object sender, EventArgs e)
        {
            string msg = "Отправлено подтверждение <ACK> (06h) по кнопке."; //ACK = "\x06";
            _serialPort.Write(ACK);   // подтверждение <ACK> string sACK = "\x06";
            Add_RTB(RTBout, $"{dtm}, {msg}.\n", Color.DarkViolet);
            if (sModes.IndexOf("Log_ACK_Button") != -1 )
            {
                WLog(msg);
            }
        }
        #endregion --- Действия по кнопкам на форме
        #region --- Дополнительные методы: GetCheckSum, ...

        private int GetCheckSum(string t)
        {
            //убираю последний символ, если это запятая
            if (t.Last().ToString() == ",")
                t = t.Substring(0, t.Length - 1);
            //ищем запятую перед контрольной суммой
            int lpos = t.LastIndexOf(",");
            if (lpos > 0) t = t.Substring(0, lpos + 1);
            //декодируем в массив байт
            byte[] array = Encoding.Default.GetBytes(t);
            //считаем сумму байт
            int sum = 0;
            foreach (byte item in array)
            {
                sum += item;
            }
            //вычитаем сумму байт строки Цельная кровь,Основной
            sum = sum - 3520;
            return sum;
        }
        #endregion --- Дополнительные методы: GetCheckSum, ...
        #region --- Тесты по кнопке Выполнить
        private void BtnRunTest_Click(object sender, EventArgs e)
        {
            MatchCollection matched;
            string testSelect = CmbTest.SelectedItem.ToString();
            switch (testSelect)
            {
                #region --- пример теста 00_ скрыт! (нажми на "+" - раскрой region :)
                case "тест 00":
                    // ...
                    Add_RTB(RTBout, $"\n {testSelect}.", Color.Blue);
                    break;
                #endregion --- пример теста 00_ скрыт! (нажми на "+" - раскрой region :)
                case "тест 01":
                    // ...
                    
                    sResult = "123.45??/?";
                    sResult = sResult.Replace("?", "");
                    sResult = sResult.Replace("/", "");
                    Add_RTB(RTBout, $"\n {testSelect}. sResult={sResult}", Color.Blue);
                    break;
                case "тест 01-1":
                    // ...
                    // Order "6O|2|2-23636||^^^1|R||||||||||||||||||||X"
                    //        0  1 2       
                    string sHist;
                    //sHist = sField[2].Trim();
                    //sHist = "2-23636";
                    //sHist = "132-77999";
                    //sHist = "1277999";    // Нет номера истории!
                    //sHist = "1-27-7999";  // будет 27-7999
                    //sHist = "12-324asd9"; // будет 324asd9
                    sHist = "2-";         // Нет номера истории!
                    //Regex reg = new Regex(@"[0-9]{1,2}-[0-9]{1,6}");  // цифры 0-9 1 или 2 раза, затем минус, затем цифры 0-9 от 1 до 6 раз
                    //matched = reg.Matches(sHist);
                    //if (matched.Count > 0)
                    //    HistoryNumber = matched[1].Value;
                    int im = sHist.IndexOf("-");
                    if (im==-1 | im==sHist.Length-1 )
                    {
                        Add_RTB(RTBout, $"\nНет номера истории!\n", Color.Red);
                        HistoryNumber = "0";
                        // return; BugFix 2021-06-03 пишем в SQL даже если нет номера истории.
                    }
                    HistoryNumber = sHist.Substring(im+1); // после первого минуса и до конца строки должен быть номер истории. 2021-05-31.
                    if (HistoryNumber.Length == 0)
                        HistoryNumber = "0";

                    Add_RTB(RTBout, $"\n {testSelect}.", Color.Blue);
                    break;
                case "тест 01-2":  // 2021-04-05
                    string inputString = "\u0005\r\u00021H|\\^&||ьBiOLiS NEO^SYSTEM1||ь|ьHOST^P_1||P|1|20210401160604\r\r\u00022P|1|2104010059||ь16332 ^??ї?ї?ї? ?.?. ??ї||ьU||ь|ь\r\r\u00023O|1||^1^59|^^Ю1^??ї?ї?ї^0\\^^Ю5^??.??ї^0\\^^Ю7^??ї?ї-?^0\\^^Ю8^??ї?ї?ї?^0\\^^Ю9^??ї^0\\^^Ю10^??ї^0\\^^Ю11^??ї?ї?ї^0\\^^Ю12^??ї.??ї^0\\^^Ю13^??-??ї^0\\^^Ю14^??ї^0\\^^Ю15^??ї.??ї.^0\\^^Ю16^??ї.??^0\\^^Ю17^??ї.??^0\\^^Ю18^?-??^0|R||ь|ь|ь|ь|Serum||ь|ь|ь|ь|F\r\u00024\r\r\u00025R|1|^^Ю1^??ї?ї?ї^0| 1.1529|??ї?ї/?| 4.200° TO  6.400°|L||F||ь|20210401063411\r\r\u00026R|2|^^Ю5^??.??ї^0|56.1505|?/?|64.00°0 TO 83.00°0|L||F||ь|20210401063210\r\r\u00027R|3|^^Ю7^??ї?ї-?^0|27.0224|??ї?ї?/?|71.00°0 TO 115.00°0|L||F||ь|20210401063255\r\r\u00020R|4|^^Ю8^??ї?ї?ї?^0|41.3238|??ї?ї/?| 1.700° TO  8.300°|H||F||ь|20210401063109\r\r\u00021R|5|^^Ю9^??ї^0|77.6034|??/?| 0.00°0 TO 40.00°0|H||F||ь|20210401063240\r\r\u00022R|6|^^Ю10^??ї^0|195.5502|??/?| 0.00°0 TO 37.00°0|H||F||ь|20210401063225\r\r\u00023R|7|^^Ю11^??ї?ї?ї^0|78.9669|??/?|28.00°0 TO 100.00°0|N||F||ь|20210401063124\r\r\u00024R|8|^^Ю12^??ї.??ї^0|215.5272|??/?|24.00°0 TO 195.00°0|H||F||ь|20210401063341\r\r\u00025R|9|^^Ю13^??-??ї^0| 7.8414|??/?| 0.00°0 TO 25.00°0|N||F||ь|20210401063356\r\r\u00026R|10|^^Ю14^??ї^0|1012.6040|??/?|230.00°0 TO 460.00°0|H||F||ь|20210401063326\r\r\u00027R|11|^^Ю15^??ї.??ї.^0|56.9256|??/?|98.00°0 TO 279.00°0|L||F||ь|20210401063310\r\r\u00020R|12|^^Ю16^??ї.??^0|25.1580|??ї?ї?/?| 0.00°0 TO 21.00°0|H||F||ь|20210401063139\r\r\u00021R|13|^^Ю17^??ї.??^0| 6.9215|??ї?ї?/?| 0.00°0 TO  6.800°|H||F||ь|20210401063155\r\r\u00022R|14|^^Ю18^?-??^0| 5.7643|??/?| 0.00°0 TO  5.00°0|H||F||ь|20210401063426\r\r\u00023L|1|N\r\r\u0004";
                    inputString += CR + LF + STX + "4";
                    //inputString = Regex.Replace(inputString, ETB + "..", String.Empty);  
                    //inputString = "Serum||||||||||FD3"+CR+LF+ "4";
                    Add_RTB(RTBout, $"\n Тест 01:\n{inputString}", Color.Green);
                    inputString = Regex.Replace(inputString, ETB + ".." + CR + LF + STX + ".", String.Empty);
                    /* это замена управляющих символов, если данные разбиты на несколько фреймов: 
                         <STX> [FN] [TEXT] <ETB> [C1][C2] <CR> <LF> Intermediate frame
                         <STX> [FN] [TEXT] <ETX> [C1][C2] <CR> <LF> Termination frame
                      Вот пример:
                        [STX]3O|1|01040218001|^1^1|^^^1^T-BIL^0・・・¥^^^18^CPK^0¥^^^19[ETB]??[CR][LF]
                        [STX]4^AMY^0|R||||||N||||Serum||||||||||F[CR][ETX]??[CR][LF]
                    */
                    inputString = Regex.Replace(inputString, CR + LF, String.Empty);
                    Add_RTB(RTBout, $"\n После замены:\n{inputString}", Color.Blue);
                    break;
                case "тест 02":    // 2020-MM-DD
                    string findkod = "";
                    string sKod = "^^^1^???????^0";  // "^^^12^???????^0"
                    Regex rg = new Regex(@"\d+");    // найти первые n цифр: Regex(@"\d+"); первые две цифры: Regex(@"\d{1,2}");
                    //MatchCollection matched = rg.Matches(sKod);
                    matched = rg.Matches(sKod);
                    if (matched.Count > 0)
                        findkod = matched[0].Value.ToString();
                    Add_RTB(RTBout, $"\n {testSelect}, kod={findkod}.", Color.Blue);
                    break;
                case "UTF32":  
                    _serialPort.Encoding = Encoding.UTF32; 
                    Stat1.Text = _serialPort.Encoding.ToString();
                    WLog(testSelect);
                    Add_RTB(RTBout, $"\n{testSelect}", Color.Blue);
                    break;
                case "UTF8": 
                    _serialPort.Encoding = Encoding.UTF8;
                    Stat1.Text = _serialPort.Encoding.ToString();
                    WLog(testSelect);
                    Add_RTB(RTBout, $"\n{testSelect}", Color.Blue);
                    break;
                case "Unicode":
                    _serialPort.Encoding = Encoding.Unicode; Stat1.Text = _serialPort.Encoding.ToString(); WLog(testSelect); Add_RTB(RTBout, $"\n{testSelect}", Color.Blue);
                    break;
                case "ASCII": 
                    _serialPort.Encoding = Encoding.ASCII; Stat1.Text = _serialPort.Encoding.ToString(); WLog(testSelect); Add_RTB(RTBout, $"\n{testSelect}", Color.Blue);
                    break;
                case "866":   
                    _serialPort.Encoding = Encoding.GetEncoding(866); Stat1.Text = _serialPort.Encoding.ToString(); WLog(testSelect); Add_RTB(RTBout, $"\n{testSelect}", Color.Blue);
                    break;
                case "1251": 
                    _serialPort.Encoding = Encoding.GetEncoding(1251); Stat1.Text = _serialPort.Encoding.ToString(); WLog(testSelect); Add_RTB(RTBout, $"\n{testSelect}", Color.Blue);
                    break;
                case "Q Encoding":    // 2020-04-20 проверка кодировки 2020-04-20
                    //qkreq= _serialPort.Encoding.ToString();
                    qkreq = _serialPort.Encoding.CodePage.ToString();
                    Stat1.Text =$"Q Encoding: {qkreq}";
                    WLog(Stat1.Text);
                    Add_RTB(RTBout, $"\n {Stat1.Text} ", Color.Blue);
                    break;
                case "перечитать ini-файл": 
                    ReadParmsIni();
                    break;
                case "Pic visible": // Pic vivible
                    Pic1_Click(sender, e);
                    break;
                default:
                    MessageBox.Show($"Нет теста для:  {CmbTest.SelectedItem}\n");
                    break;
            }
            // --- end switch
        }
        #endregion --- Тесты по кнопке Выполнить
    }
}
#region --- Automated Clinical Analyzer. BiOLiS 24i Premium. Bi-directional Communication Specifications. Version 1.07.
/* 
 3.Protocol of Data Link Layer
 The protocol of data link layer is defined by ASTM 1381-91.
 The protocol of data link layer uses the following transmission control codes.
 No. Transmission control code name    
      |    Transmission control code
      |     |
      V     V          Explanation
  1  <STX>  2(02h)   Code to show the beginning of text. 
  2  <ETB> 23(17h)   Code to show the interruption of text. When the transmitted text is too large, it is split into multiple frames, using <ETB>.
  3  <ETX>  3(03h)   Code to show the end of text. 
  4  <CR>  13(0Dh)   Carriage return
  5  <LF>  10(0Ah)   Line feed code 
  6  <ENQ>  5(05h)   Enquiry
  7  <ACK>  6(06h)   Acknowledge
  8  <NAK> 21(15h)   Not acknowledge 
  9  <EOT>  4(04h)   End of transmission 
 10  [FN]     -      Frame number.  ASCII numbers from 0 to 7. The first frame begins with 1. 
 11  [C1][C2] -      Checksum

 3.1.Frame

  <STX> [FN] [TEXT] <ETB> [C1] [C2] <CR> <LF> Intermediate frame
  <STX> [FN] [TEXT] <ETX> [C1] [C2] <CR> <LF> Termination frame

  [STX]3O|1|01040218001|^1^1|^^^1^T-BIL^0・・・¥^^^18^CPK^0¥^^^19[ETB]??[CR][LF]
  [STX]4^AMY^0|R||||||N||||Serum||||||||||F[CR][ETX]??[CR][LF]

 Note: All the example of check sum value is invalid. 
 [TEXT] is the text data to be transmitted. In machine, [TEXT] corresponds to the record (ASTM 1394-91).  
 230 characters are the maximum in [TEXT].   (230 octets)   
 The text, which exceeds 230 octets, uses multiple frames, using < ETB >.
 Checksum is the least 8 bits of the value, that is gotten when the sum of character codes from [FN] to <ETB>,< ETX >. (Modulo 256).
 [C1] and [C2] are ASCII alphanumeric hexadecimal notations of the upper 4 bits and the lower 4 bits of checksum, respectively.

 4.3.  The Maximum Length of Record 
 The maximum length of record is a record that includes 1024 characters(1024 octets).
 The escape characters are counted for the characters after escape. (For example, “&F &” is counted as three characters).
 Note: One line can transmit only 230 characters. If more than 230 characters need to be transmitted, please use second lines by dividing <ETB>.

 Automated Clinical Analyzer.   BiOLiS 24i Premium.  Bi - directional Communication Specifications.   Version 1.07.   page 13
 The records are defined by ASTM1394-91. The records supported by BiOLiS are as follows. 
    No Record_ID   Record 
     1      H    Message Header Record 
     2      P    Patient Information Record 
     3      O    Measurement Order Record 
     4      Q    Enquiry Record 
     5      C    Comment Record 
     6      R    Measurement Result Record 
     7      L    Message Terminator Record 
*/
#endregion --- Automated Clinical Analyzer. BiOLiS 24i Premium. Bi-directional Communication Specifications. Version 1.07.
#region --- Выполняемые анализы ---
/* Код  Анализ	    Ед.изм.	Референсные значения
 1	ГЛЮКОЗА     ммоль/л	4.2 – 6.4
 2	ХОЛЕСТ-Н
 3	ТРИГЛИЦ
 4	ЛПВП
 5	Общий белок	г/л	     64 – 83
 7	КРЕАТИНИН	ммоль/л	 71 – 115
 8	МОЧЕВИНА	ммоль/л	1.7 – 8.3
 9	АЛТ	Ед/л	          0 – 40
 10	АСТ	Ед/л	          0 – 37
 11	АМИЛАЗА	    Ед/л	 28 – 100
 12	КФК ОБЩ.	Ед/л	 24 – 195
 13	КФК МБ	    Ед/л	  0 – 25
 14	ЛДГ     	Ед/л	230 – 460
 15	ЩЕЛ.ФОС.	Ед/л	 98 – 279
 16	БИЛ.ОБ.	    ммоль/л	  0 – 21
 17	БИЛ.ПР		
 18	СР-Б		
 83	ЛПНП		
 90	КА		
*/
#endregion --- Выполняемые анализы ---
