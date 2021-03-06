﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
//using System.Net;
using System.Reflection;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using System.IO.Ports;
//using System.Text.RegularExpressions;
//using Excel = Microsoft.Office.Interop.Excel;
#region --- лицензия GNU
//********************************************************************************************
//                           LICENSE INFORMATION
//********************************************************************************************
// Sapphire 400 Automated Clinical Analyzer Driver, 
// serial port & SQL communications. Version 1.04.57.
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
        private static string sModes;       // строка списка режимов работы 
        private static string PathLog;      // путь к лог-файлу с вычисленными каталогом и датой ...\GGGG-MM\BeckLog2019-08-06.txt 
        private static string PathLogDir;   // путь к лог-файлу, заданный в параметрах ini-файла  
        private static string PathErrLog;   // путь к лог-файлу ошибок 
        //private string PathRep = "";        // путь к файлам с названиями отчётов.sql
        //private string PathDoc = "";        // путь к файлам с результирующими отчётами (документы .xls .doc и т.д.)
        private string PathIni;             // Путь к ini-файлу (там, откуда запуск)
        private string inputString = ""; // полученные данные - для парсинга
        // глобальные
        private static DateTime dt0 = DateTime.Now; // время старта 
        private static DateTime dtm = dt0;          // для измерения времени выполнения запроса 
        private string sTimeStart   = dt0.ToString("dd-MM-yyyy HH:mm:ss");
        private int    NumerOfStart = 0;
        private string UserName = System.Environment.UserName;
        private string ComputerName = System.Environment.MachineName;
        //private string myIP="" ;
        private string strVer = "первоначальная версия - 1.0.1. ( потом изменю :)"; // :))
        private static string AppName;      // static т.к. исп. в background-процессе
        private static string qkreq = "";   // запрос о работе ("кукареку")
        private static string sHeader1 = "";// строка заголовка - назначение, для чего.
        private string msg, msg0,msg1,msg2;
        private string ComPortName;
        private string ENQ = "\x05", ACK = "\x06", NAK = "\x15", ETB = "\x17" 
                      ,STX = "\x02", ETX = "\x03", EOT = "\x04", CR  = "\n", LF = "\r"; // управляющие коды в протоколе обмена для СОМ-порта
        private Random rnd = new Random(3213);
        private int partNo = 0;   // i-я часть принятых данных
        #endregion --- Общие параметры
        #region --- для SQL
        // для результата Select
        private string scriptSql    = "";  // текст Select из файла  - file.OpenText().ReadToEnd();
        private string scriptSqlPar = "";  // текст с параметрами из файла со скриптом (Select...)
        private static string dateDone = "2019-12-31 23:59";  // дата-время выполнения анализа по часам на анализаторе
        private static string dateDone999 = "31-12-2019 23:59:00.000";  // дата-время выполнения анализа для SQL
        //private string st1_0 = "Insert into AnalyzerResults (Analyzer_id, HostName, HistoryNumber, ResultDate, CntParam, ResultText";
        private string st1_0 = "Insert into AnalyzerResults (Analyzer_id, HostName, HistoryNumber, ResultDate, CntParam";
        private string st1_1 = "";
        //private string st2_0 = $") values ({Analyzer_Id}, host_name(), {HistoryNumber}, {ResultDate}, {CntParam}, {ResultText}";    // для строки Insert...
        private string st2_0 = ""; //$") values ({Analyzer_Id}, host_name(), {HistoryNumber}, GetDate(), {CntParam}, {ResultText}";    // для строки Insert...
        private string st2_1 = "";                          // для строки Insert...
        private string st0 = "";
        private static string HistoryNumber = "0", sResult = "", sEdizm = "", sRefVal = "", sFlag = "", dt_Done = "", ResultDate="", ResultText="";
        private static int    CntParam;
        private static int    MaxCntParam;       // = 22 на 2020-04-24
        private static int    MaxSapphireResult; // = 12 на 2020-04-24
        private static bool   IsControl = false; // признак - текущая принятая запись - это контроли.
        private string        patternHistN = ""; // для выделения HistoryNumber  @"\d{0,6}";  - от 0 до 6 цифр максимум! 
        private static int    nHistNo = -1;
        private int kRecord = 0; // кол-во полученных записей
        private int kRes = 0;
        //private int kNpp = 0; // кол-во постановок, образцов, N по порядку, как в бумажном журнале...
        #endregion --- для SQL
        public FormSap()
        {
            InitializeComponent();
            ReadParmsIni();         // берём параметры из ini-файла 
            FormIni();              // установки на форме, которые не делает Visual Studio - IP-адреса          
            PortIni();
        }
        private void FormIni()          // мой начальный вывод на форму, что прочитали из ini-файла
        {
            CmbTest.SelectedIndex = 0;  // первый (нулевой) элемент - текущий, видимый.
            this.Pic1.Image = new Bitmap($"{PathIni}\\Pic.png");  // на форме картинка - для различных приложений должна быть другая!
        }
        // **************************************
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
            //Pic_Status.Visible = true; ### сделать изменяемую картинку в статусной строке
            dtm = DateTime.Now;
            if (dt0.ToString("yyyy-MM-dd") != dtm.ToString("yyyy-MM-dd"))
            {
                dt0 = dtm;
                SetPathLog(); // изменить какалог при смене даты
            }

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
            if (iEOT != -1) // есть признак конца <EOT> 04h ((kb == 1) & (lastByte == EOT) )
            {
                Add_RTB(RTBout, $"{dtm} есть <EOT>\n", Color.DarkGreen);
                if (kb != 1)    // должен приходить 1 байт
                {
                    msg = "ERR: получен признак <EOT> и ещё данные, они не обработаны! :(";
                    Add_RTB(RTBout, $"{dtm} {msg}\n", Color.Red);
                    //MessageBox.Show(msg, "Внимание! Ошибка в двнных!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    msg += $"\n partNo={partNo}\n iEOT={iEOT}\n После <EOT>:{receivedData.Substring(iEOT)}\n receivedData={receivedData}";
                    WErrLog(msg);
                    WLog(msg);
                }

                ParseResults(inputString);  // вызвать обработчик ...

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
                else if (CntParam > 22) // возможно, это повторы (2020-04-16) 
                {
                    // записать собщение для лаборанта...
                    msg = "Больше 22 анализов! Это повторы: их пока пропускаем (22 поля для приёма в LabAutoResult!!)";
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
            #region  --- Описание типов полей(записей) в протокое обмена ASTM1394-91
            /* Automated Clinical Analyzer.   BiOLiS 24i Premium.  Bi - directional Communication Specifications.   Version 1.07.
             *     page 13
             * The records are defined by ASTM1394-91. The records supported by BiOLiS are as follows. 
             *   No Record_ID   Record 
             *    1      H    Message Header Record 
             *    2      P    Patient Information Record 
             *    3      O    Measurement Order Record 
             *    4      Q    Enquiry Record 
             *    5      C    Comment Record 
             *    6      R    Measurement Result Record 
             *    7      L    Message Terminator Record 
            */
            #endregion  --- Описание типов полей в протокое обмена ASTM1394-91
            Stat1.Text = "парсинг";
            Stat1.ForeColor = Color.DarkBlue;
            dtm = DateTime.Now;
            Add_RTB(RTBout, $"{dtm} Начало парсинга, {inputString.Count()} байт получено для парсинга:\n", Color.DarkViolet);
            Add_RTB(RTBout, $"{inputString}", Color.Black);

            IsControl = false;
            st1_1 = ""; st2_1 = "";
            CntParam = 0;
            //  удаление ETX и следующих за ним двух байт контрольной суммы - они не нужны мне.
            inputString = Regex.Replace(inputString, ETX + "..", String.Empty);
            inputString = Regex.Replace(inputString, CR + LF, String.Empty);    // удаление CR+LF

            //2020-04-24 - сделать в параметрах ini-файла! 
            //string patternHistN = @"\d{0,6}";   // для выделения HistoryNumber  2020-04-24 - сделать в параметрах ini-файла! 
            Regex rg = new Regex(patternHistN);
            MatchCollection matched;
            string[] sDigits;
            string sHist = "";  // " 25781/ ТЕСТОВЫЙ АФ АРО1";   "?16512 ^??????? ?.?. ???1"  - это первоначальная строка с номером ИБ, ФИО, ...

            string[] sRecord = inputString.Split(new string[] { CR, LF, ENQ, STX, EOT }, StringSplitOptions.RemoveEmptyEntries);
            int ks = sRecord.Count();
            for (int i = 0; i < ks; i++)  // по количеству строк в переданных результатах пациента
            {
                string[] sField = sRecord[i].Split('|');
                int kField = sField.Count();
                string sRecordType = sRecord[i].Substring(1, 1);
                if (sRecordType == "P")     // Patient  "2P|1|2020040150101|||19576 ^??????? ?.?   ???|||U|||||"
                {                           //           0  1 2            345                        67 
                    if (kField == 2)
                    {
                        IsControl = true;
                        Add_RTB(RTBout, $"\n - Это контроли, их пропускаем...", Color.DarkGreen);
                        break;  // это  контроли, их пока не обрабатываем (но надо, я их в Науке собирал для внутрилабораторных отчётов)
                    }
                    sHist = sField[5].Trim();
                    matched = rg.Matches(sHist);
                    sDigits = Regex.Split(sHist, patternHistN);
                    HistoryNumber = "0";
                    for (int j = 0; j < matched.Count; j++)
                    {
                        if (matched[j].Value.Length != 0)
                        {
                            HistoryNumber = matched[j].Value;
                            break;  // результат д.б. в первой непустой
                        }
                    }
                }

                if (sRecordType == "R")   // Result "4R|1|^^^1^???????^0| 4.7804|?????/?| 4.2000 TO  6.4000|N||F||||20200415165509"
                {                         //         0  1 2             3        4       5                  6 78 9012 

                    string sKod = sField[2].Substring(3, 2);
                    sKod = Regex.Replace(sKod, "\\^", String.Empty);    // удалить строку "^"
                    sResult = sField[3].Trim();
                    sEdizm = sField[4].Trim();
                    sRefVal = sField[5].Trim();
                    sRefVal = Regex.Replace(sRefVal, @"\s+", " ");       // заменяет несколько подряд идущих пробелов одинарными
                    if (sRefVal.Length > 16)
                        sRefVal = "длина больше 16";
                    sFlag = sField[6];

                    if (kField >= 13)   // для контроля индекса dt_Done = sField[12];
                        dt_Done = sField[12];
                    else if (kField == 12)
                        dt_Done = sField[11];
                    else //  (kField < 12)                   
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

                    st1_1 += $", ParamName{CntParam}, ParamValue{CntParam}, ParamRef{CntParam}";    // + ParamMgr{CntParam}
                    st2_1 += $", '{sKod}', '{sResult}', '{sRefVal}'";
                    //Add_RTB(RTBout, $"\nКод:{sKod}, рез.:{sResult}, Ед.изм:{sEdizm}, реф:{sRefVal}, Вр.вып:{dt_Done}.", Color.Brown);
                    Add_RTB(RTBout, $"\nКод:{sKod}, результат:{sResult}, реф.знач:{sRefVal}, выполнен:{dt_Done}.", Color.DeepPink);
                }
            }

            // формирование строки SQL Insert...
            ResultText = inputString;
            //ResultDate= Get Date()
            //private string st1_0 = "Insert into AnalyzerResults (Analyzer_id, HostName, HistoryNumber, ResultDate, CntParam";
            //st2_0 = $") values ({Analyzer_Id}, host_name(), {HistoryNumber}, GetDate(), {CntParam}, '{ResultText}' ";    // для строки Insert...
            st2_0 = $", ResultText) values ({Analyzer_Id}, host_name(), {HistoryNumber}, GetDate(), {CntParam} ";    // для строки Insert...
            st0 = st1_0 + st1_1 + st2_0 + st2_1 + $", '{ResultText}');";
            // ResultText 
            //Add_RTB(RTBout, $" SQL:\n{st0}",Color.Red);

            dtm = DateTime.Now;
            Add_RTB(RTBout, $"\n{dtm} Конец парсинга [{kRecord++}]. HistoryNumber:{HistoryNumber}.\n", Color.DarkViolet);
            WLog($"Получено для парсинга {inputString.Count()} байт:\n{inputString}");
            Stat1.Text = "конец парсинга";
            //Stat1.ForeColor = Color.DarkBlue;
            #region --- old parse
            /*
            foreach (string s in text)
            {
                kNpp++;
                string preparedS = s;
                if (s.Length > 13)
                {
                    preparedS = s.Remove(0, 13);
                    //textBoxResult.AppendText("Prepared (without first 14 symbols): " + preparedS + "\n");
                    //try
                    //{
                    Match res = Regex.Match(preparedS, parsePattern);
                    cData = res.Groups["Data"].Value;
                    NHist = res.Groups["HistNumber"].Value;
                    NSamp = res.Groups["SampNumber"].Value;
                    kRes = 0;

                    // 2019-08-09 раскраска текста RichTextBox :)))
                    Add_RTB(RTBout, $"\n Начало kNpp {kNpp}, NHist {NHist}, NSamp {NSamp} \n", Color.Red);
                    Add_RTB(RTBout, $"{cData}", Color.Black);

                   
                    // очистить список результатов
                    // hres.Clear();
                    cresAll = ""; // разобранные результаты для лог-файла
                    stNames = ""; // заготовки для Insert - для каждого пациента
                    stValues = ""; // заготовки для Insert - для каждого пациента
                    string stData = res.Groups["Data"].Value;

                    foreach (Match dataPair in Regex.Matches(res.Groups["Data"].Value, dataParsePattern))
                    {
                        kRes++;
                        ckod = dataPair.Groups["SampleID"].Value;
                        cres = dataPair.Groups["SampleResult"].Value;
                        cires = dataPair.Groups["SampleLetters"].Value;

                        Add_RTB(RTBout, $"\n ckod {ckod} ===== \n cres {cres}\n cires-{cires}-", Color.DarkGreen);

                        // формируем конец строки Insert: ParamName1,ParamValue1,ParamName2,...
                        //if (sendIRes == "2") //  только пара <код_анализа> <результат>
                        if ( "2" == "2") //  только пара <код_анализа> <результат> // заглушка
                        {
                            stNames += $",ParamName{kRes},ParamValue{kRes}";
                            stValues += $",'{ckod}','{cres}'";
                            cresAll += $"[{kRes}]{ckod};{cres}";
                        }
                        else //  тройка <код_анализа> <результат> <интерпретация_результата>
                        {
                            stNames += $",ParamName{kRes},ParamValue{kRes},ParamMsr{kRes}";
                            stValues += $",'{ckod}','{cres}','{cires}'";
                            cresAll += $"[{kRes}]{ckod};{cres};{cires}";
                        }

                    }

                    Add_RTB(RTBout, $"\n Конец. kNpp {kNpp}, NHist {NHist}, NSamp {NSamp}, " +
                        $"kRes {kRes} {dtm}", Color.Red);


                    WLog($"NHist {NHist}, NSamp {NSamp}, kRes {kRes}. {cresAll}"); //2019-08-13 kRes {kRes}.

                    // формирование строки Insert...
                    st1_1 = ", ResultDate, HistoryNumber, CntParam, ResultText";
                    string ResText = $"{NHist} {NSamp} {stData}"; //2019-07-31
                                                                  //st2_1 = $", GetDate(), {NHist}, {kRes}, '{NHist} {NSamp} {stData}'  ";
                    st2_1 = $", GetDate(), {NHist}, {kRes}, '{ResText}' ";
                    st0 = st1_0 + st1_1 + stNames + ") " + st2_0 + st2_1 + stValues + "); ";
                    WLog("SQL: " + st0);          // в лог-файл

                    if ((NHist.Count() == 0) || (NSamp.Count() == 0))
                        WLog($"--- Нет номера истории или номера пробы! {NHist}, {NSamp}."); // и не писать в SQL
                                                                                             // проверка на уже записанные сегодня результаты по этому HistoryNumber - проверка в SQL. Но ЕСТЬ ЕЩЁ 3 ВАРИАНТА!!! 02.08.2019.
                    else if (CheckExist(ResText) > 0)
                        WLog("--- Дубль не записан.");
                    else if (true) // (chk_WriteSQL.Checked) // писать в SQL - флажок на форме? или лучше в параметрах? пока на форме!
                        ToSQL(st0);  // запись в SQL
                    else
                    {
                        // всё хорошо, ничего делать больше не надо :))
                    }
                    
                }
                
            }
            */
            #endregion --- old parse

        }
        #endregion --- Парсинг данных из строки и формирование строки SQL Insert        // ---
        private void ToSQL(string st) // запись в MS-SQL сформированной строки
        {
            //if (IsControl) // контроли пропускаем
            ///    return;

            //using (SqlConnection sqlConn = new SqlConnection(Properties.Settings.Default.connStr))
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
            PathIni = Application.StartupPath;
            AppName = AppDomain.CurrentDomain.FriendlyName;
            AppName = AppName.Substring(0, AppName.IndexOf(".exe"));
            string pathIniFile = PathIni + @"\" + $"\\{AppName}" + ".ini";
            if (!File.Exists(pathIniFile))
            {
                string errmsg = "Не найден файл " + pathIniFile+"\n Работа завершается!";
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
            connStr = iniFile.GetPrivateString("Connection", "DbSQL").Trim(); // получить значение  из секции 1.. по ключу 2..
            string nameSQLsrev = connStr.Substring(0, connStr.IndexOf(';'));
            ComPortName= iniFile.GetPrivateString("Connection", "ComPortNo").Trim();
            // секция [Check]
            /*
            [Check]
            patternHistN=\d{0,6}
            MaxCntParam=22
            MaxSapphireResult=12
            */
            patternHistN = iniFile.GetPrivateString("Check", "patternHistN").Trim(); // шаблон для выделения номера истории по RegEx patternHistN=\d{0,6}
            str = iniFile.GetPrivateString("Check", "MaxCntParam").Trim();
            MaxCntParam = Convert.ToInt32(str);
            str = iniFile.GetPrivateString("Check", "MaxSapphireResult").Trim();
            MaxSapphireResult = Convert.ToInt32(str);

            // секция [LogFiles]
            PathLogDir = iniFile.GetPrivateString("LogFiles", "PathLogDir");
            SetPathLog();
            PathErrLog = iniFile.GetPrivateString("LogFiles", "PathErrLog");

            // секция [Modes]  Режимы работы 
            sModes = iniFile.GetPrivateString("Modes", "mode1");
            /* Пример содержимого:
            mode1 = (Log_Receive),(...), ...
            */

            // секция [Comments]
            string sNumerOfStart = iniFile.GetPrivateString("Comments", "StartNo");
            NumerOfStart =Convert.ToInt32(sNumerOfStart);
            NumerOfStart++;
            sNumerOfStart = Convert.ToString(NumerOfStart);
            iniFile.WritePrivateString("Comments", "StartNo", sNumerOfStart );
            iniFile.WritePrivateString("Comments", "LastStartTime", dtm.ToString() );// записать значение в секции Comments по ключу xx

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            //var vers = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            //var fileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).FileVersion;
            //var productVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).ProductVersion;
            //string sCreationTime  = new System.IO.FileInfo(assemblyLocation).CreationTime.ToString();
            string sLastWriteTime = new System.IO.FileInfo(assemblyLocation).LastWriteTime.ToString();
            strVer = $"Версия {version} от {sLastWriteTime}"; //versionString;
            WLog($"--- Запуск {AppName} {ComputerName} {UserName} {strVer}");
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
            FileStream fn = new FileStream(PathLog, FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
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
        private static void WTest(string FileNam, string st) // записать в FileNam.txt
        {
            FileStream fn = new FileStream(PathErrLog + "\\" + FileNam + ".txt", FileMode.Append);
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
        private void Form1_FormClosing(object sender, FormClosingEventArgs e) // закрыть App
        {
            // Закрытие формы - FormClosing
            msg = "Завершение работы.";
            DialogResult result = MessageBox.Show("Вы действительно хотите завершить работу\n c программой?"
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
            //ExitApp(mess);
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
            //    WLog("--- передумал выходить :) ");
            //    return; // просто передумал :)
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
            s += $"         {sHeader1}\n";
            s += $"AppName: {AppName}\n";
            s += $"{strVer}\n";
            s += sep;
            s += $"Время старта: {sTimeStart}\n";
            s += $"ComputerName: {ComputerName}, UserName: {UserName}\n";
            //s += $"IP: {myIP}\n";
            s += sep;
            s += $"путь к логам:\n";
            s += $"PathLog: {PathLog}\n";
            s += $"PathLogDir: {PathLogDir}\n";
            s += $"PathIni: {PathIni}\n";
            s += sep;
            s += $"Режимы работы: {sModes}\n";
            s += $"SQL: {nameSQLsrv}, Analyzer_Id: {Analyzer_Id}\n"; 
            s += sep+"\n\n\n\n";
            DialogResult result = MessageBox.Show(s, "  Параметры в .ini-файле:"
                , MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExitApp("Выход по меню <Выход>");
        }
        private void ВыходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ExitApp("Выход из меню <Выход>");
        }
        #endregion --- методы Wlog, WErrLog;  SetPathLog, Add_RTB, ExitApp...
        #region --- Действия по кнопкам на форме
        private void Btn_ini_Click(object sender, EventArgs e)  // Кнопка: прочитать ещё раз ini-файл
        {
            ReadParmsIni();
        }
        private void Btn_Clear_Click(object sender, EventArgs e) // Кнопка: очистить форму
        {
            RTBout.Clear();
            partNo = 0;
            kRecord = 0;
        }
        private void BtnRun_Click(object sender, EventArgs e)   // кнопка Выполнить отчёт с параметрами
        {
            // кнопка Выполнить отчёт с параметрами
            //FormParmRep formParmRep = new FormParmRep(ParmRep.LastRep);
            //formParmRep.ShowDialog();   // здесь заполнили параметры выбранного отчёта

            MessageBox.Show($"Кто послал: {sender},\n аргументы: {e}.\n", "Выполнить отчёт с параметрами (BtnRun_Click)."
                , MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        #region --- Дополнительные методы: FormMain, FormResize, GetCheckSum, ...
        /*
        public FormMain()
        {
            //проверка на посторный запуск
            int prC = 0;
            foreach (Process pr in Process.GetProcesses())
                if (pr.ProcessName == Process.GetCurrentProcess().ProcessName) prC++;
            if (prC > 1) Process.GetCurrentProcess().Kill();
            InitializeComponent();
            // делаем невидимой нашу иконку в трее
            notifyIcon1.Visible = false;

            // добавляем Эвент или событие по 2му клику мышки, 

            //вызывая функцию  notifyIcon1_MouseDoubleClick
            this.notifyIcon1.MouseDoubleClick += new MouseEventHandler(notifyIcon1_MouseDoubleClick);

            // добавляем событие на изменение окна
            this.Resize += new System.EventHandler(this.Form1_Resize);

        }
        */

        /*
        private void FormResize(object sender, EventArgs e)
        {
            if (ToTray)
            {
                // проверяем наше окно, и если оно было свернуто, делаем событие        
                if (WindowState == FormWindowState.Minimized)
                {
                    // прячем наше окно из панели
                    this.ShowInTaskbar = false;
                    // делаем нашу иконку в трее активной
                    notifyIcon1.Visible = true;
                }
            }
        }
        */

        /*
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (ToTray)
            {
                // делаем нашу иконку скрытой
                notifyIcon1.Visible = false;
                // возвращаем отображение окна в панели
                this.ShowInTaskbar = true;
                //разворачиваем окно
                WindowState = FormWindowState.Normal;
            }
        }
        */
        /// <summary>
        /// Считает контрольную сумму
        /// </summary>
        /// <param name="t">Строка для подсчека контрольной суммы</param>
        /// <returns>Сумму байт строки</returns>
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
            string testSelect = CmbTest.SelectedItem.ToString();
            switch (testSelect)
            {
                #region --- пример теста 00_ скрыт! (нажми на "+" - раскрой region :)
                case "тест 01":    // 2020-MM-DD
                    // ...
                    MessageBox.Show($"{scriptSql}\n *** Параметры:\n{scriptSqlPar}\n ***", "Тест (пример) 01:"
                        , MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    break;
                #endregion --- пример теста 00_ скрыт! (нажми на "+" - раскрой region :)
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
                #region --- done cases or not allready used
                case "тест Выделение HistoryNumber":    // 2020-04-13
                    string patternHistN = @"\D+";
                    patternHistN = @"\d{0,6}";
                    Regex rg = new Regex(patternHistN);
                    MatchCollection matched;

                    //string[] sDigits = Regex.Split(sText, @"\D+");
                    //Regex.Replace(inputString, pattern, String.Empty);
                    string[] sDigits;
                    Add_RTB(RTBout, "\n---START---", Color.Blue);

                    string sTest = " 25781/ ТЕСТОВЫЙ АФ АРО1";
                    sTest = sTest.Trim();
                    matched = rg.Matches(sTest);
                    sDigits = Regex.Split(sTest, patternHistN);
                    HistoryNumber = "0" + matched[0].Value;
                    //HistoryNumber = "0"+Regex.Replace(sTest, sPattern, String.Empty); // удалить все не-цифры
                    //HistoryNumber = "0"+sDigits[0];
                    Add_RTB(RTBout, $"\n 1 sTest:{sTest}, HistoryNumber:{HistoryNumber}.", Color.Blue);

                    sTest = "9999/ ТРЕВОЖНАЯ КН АРО5";
                    sTest = sTest.Trim();
                    matched = rg.Matches(sTest);
                    sDigits = Regex.Split(sTest, patternHistN);
                    HistoryNumber = "0" + matched[0].Value;
                    //HistoryNumber = "0" + Regex.Replace(sTest, sPattern, String.Empty); // удалить все не-цифры
                    Add_RTB(RTBout, $"\n 2 sTest:{sTest}, HistoryNumber:{HistoryNumber}.", Color.Blue);

                    sTest = "????? ШУТНИК О.Н, РХМДЛ ";
                    sTest = sTest.Trim();
                    matched = rg.Matches(sTest);
                    sDigits = Regex.Split(sTest, patternHistN);
                    HistoryNumber="0"+matched[0].Value;
                    //HistoryNumber = "0" + Regex.Replace(sTest, sPattern, String.Empty); // удалить все не-цифры
                    Add_RTB(RTBout, $"\n 3 sTest:{sTest}, HistoryNumber:{HistoryNumber}.", Color.Blue);

                    sTest = "????? ШУТНИК О.Н, АРО2 ";
                    sTest = sTest.Trim();
                    matched = rg.Matches(sTest);
                    sDigits = Regex.Split(sTest, patternHistN);
                    HistoryNumber = "0" + matched[0].Value;
                    Add_RTB(RTBout, $"\n 4 sTest:{sTest}, HistoryNumber:{HistoryNumber}.", Color.Blue);


                    //  .Split( ... ,StringSplitOptions.RemoveEmptyEntries);
                    /*int kSD = sDigits.Count();
                    Add_RTB(RTBout, $"\n{kSD} - строк в sDigits", Color.Green);
                    for ( int j=0; j<kSD; j++)
                    {
                        Add_RTB(RTBout, $"\n --- sDigits[{j}]: {sDigits[j]}", Color.Blue);
                        int number;
                        if (int.TryParse(sDigits[j], out number))
                        {
                            //Console.WriteLine(value);
                            Add_RTB(RTBout, $"\n{j}-й: {sDigits[j]}", Color.Black);
                        }
                    }*/
                    Add_RTB(RTBout, "\n---END---", Color.Blue);
                    // ---...===
                    MessageBox.Show($"=== HistoryNumber===", "Тест (пример) 01:"
                        , MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    break;
                #region --- тест inputString
                case "тест inputString":    // 2020-02-05
                                   //string sMes = "тест inputString";
                                   //string pattern = @"\b[M]\w+";

                    //inputString = "\u0005\u00021H|\\^&|||BiOLiS NEO^SYSTEM1|||||HOST^P_1||P|1|20200407122247\r\u0003EB\r\n\u00022P|1|202004070025|||56174 ^??????? ?.?  ???|||U|||||\r\u0003C1\r\n\u00023O|1||^1^25|^^^1^???????^0\\^^^5^??.???^0\\^^^7^?????-?^0\\^^^8^????????^0\\^^^9^???^0\\^^^10^???^0\\^^^11^???????^0\\^^^12^???.???^0\\^^^13^???-??^0\\^^^14^???^0\\^^^15^???.???.^0\\^^^16^???.??^0|R||||||||||Serum||||||||||F\r\u0003DD\r\n\u00024R|1|^^^1^???????^0| 5.8558|?????/?| 4.2000 TO  6.4000|N||F||||20200407122232\r\u00034B\r\n\u00025R|2|^^^5^??.???^0|47.4037|?/?|64.0000 TO 83.0000|L||F||||20200407122031\r\u00032F\r\n\u00026R|3|^^^7^?????-?^0|50.3494|??????/?|71.0000 TO 115.0000|L||F||||20200407122116\r\u0003DA\r\n\u00027R|4|^^^8^????????^0| 3.3998|?????/?| 1.7000 TO  8.3000|N||F||||20200407121945\r\u0003A5\r\n\u00020R|5|^^^9^???^0|26.0191|??/?| 0.0000 TO 40.0000|N||F||||20200407122101\r\u00039D\r\n\u00021R|6|^^^10^???^0|57.2710|??/?| 0.0000 TO 37.0000|H||F||||20200407122046\r\u0003D2\r\n\u00022R|7|^^^11^???????^0|15.5304|??/?|28.0000 TO 100.0000|L||F||||20200407122000\r\u000308\r\n\u00023R|8|^^^12^???.???^0|180.3266|??/?|24.0000 TO 195.0000|N||F||||20200407122202\r\u000342\r\n\u00024R|9|^^^13^???-??^0|46.7134|??/?| 0.0000 TO 25.0000|H||F||||20200407122217\r\u000386\r\n\u00025R|10|^^^14^???^0|709.1733|??/?|230.0000 TO 460.0000|H||F||||20200407122147\r\u0003B4\r\n\u00026R|11|^^^15^???.???.^0|488.5570|??/?|98.0000 TO 279.0000|H||F||||20200407122131\r\u0003B4\r\n\u00027R|12|^^^16^???.??^0|11.7090|??????/?| 0.0000 TO 21.0000|N||F||||20200407122016\r\u0003AB\r\n\u00020L|1|N\r\u000303\r\n\u0004";

                    inputString = "\u0005\u0002" +
                        "1H|\\^&|||BiOLiS NEO^SYSTEM1|||||HOST^P_1||P|1|20200407122247\r\u0003EB\r\n\u0002" +
                        "2P|1|202004070025|||56174 ^??????? ?.?  ???|||U|||||\r\u0003C1\r\n\u0002" +
                        "3O|1||^1^25|^^^1^???????^0\\^^^5^??.???^0\\^^^7^?????-?^0\\^^^8^????????^0\\^^^9^???^0\\^^^10^???^0\\^^^11^???????^0\\^^^12^???.???^0\\^^^13^???-??^0\\^^^14^???^0\\^^^15^???.???.^0\\^^^16^???.??^0|R||||||||||Serum||||||||||F\r\u0003DD\r\n\u0002" +
                        "4R|1|^^^1^???????^0| 5.8558|?????/?| 4.2000 TO  6.4000|N||F||||20200407122232\r\u00034B\r\n\u0002" +
                        "5R|2|^^^5^??.???^0|47.4037|?/?|64.0000 TO 83.0000|L||F||||20200407122031\r\u00032F\r\n\u0002" +
                        "6R|3|^^^7^?????-?^0|50.3494|??????/?|71.0000 TO 115.0000|L||F||||20200407122116\r\u0003DA\r\n\u0002" +
                        "7R|4|^^^8^????????^0| 3.3998|?????/?| 1.7000 TO  8.3000|N||F||||20200407121945\r\u0003A5\r\n\u0002" +
                        "0R|5|^^^9^???^0|26.0191|??/?| 0.0000 TO 40.0000|N||F||||20200407122101\r\u00039D\r\n\u0002" +
                        "1R|6|^^^10^???^0|57.2710|??/?| 0.0000 TO 37.0000|H||F||||20200407122046\r\u0003D2\r\n\u0002" +
                        "2R|7|^^^11^???????^0|15.5304|??/?|28.0000 TO 100.0000|L||F||||20200407122000\r\u000308\r\n\u0002" +
                        "3R|8|^^^12^???.???^0|180.3266|??/?|24.0000 TO 195.0000|N||F||||20200407122202\r\u000342\r\n\u0002" +
                        "4R|9|^^^13^???-??^0|46.7134|??/?| 0.0000 TO 25.0000|H||F||||20200407122217\r\u000386\r\n\u0002" +
                        "5R|10|^^^14^???^0|709.1733|??/?|230.0000 TO 460.0000|H||F||||20200407122147\r\u0003B4\r\n\u0002" +
                        "6R|11|^^^15^???.???.^0|488.5570|??/?|98.0000 TO 279.0000|H||F||||20200407122131\r\u0003B4\r\n\u0002" +
                        "7R|12|^^^16^???.??^0|11.7090|??????/?| 0.0000 TO 21.0000|N||F||||20200407122016\r\u0003AB\r\n\u0002" +
                        "0L|1|N\r\u000303\r\n\u0004"; 
                    /*inputString = "\u0005\u0002" +
                    "1H|\\^&|||BiOLiS NEO^SYSTEM1|||||HOST^P_1||P|1|20200407122247\r\u0003EB\r\n\u0002" +
                    "2P|1|202004070025|||56174 ^??????? ?.?  ???|||U|||||\r\u0003C1\r\n\u0002" +
                    "3O|1||^1^25|^^^1^???????^0\\^^^5^??.???^0\\^^^7^?????-?^0\\^^^8^????????^0\\^^^9^???^0\\^^^10^???^0\\^^^11^???????^0\\^^^12^???.???^0\\^^^13^???-??^0\\^^^14^???^0\\^^^15^???.???.^0\\^^^16^???.??^0|R||||||||||Serum||||||||||F\r\u0003DD\r\n\u0002" +
                    "4R|1|^^^1^???????^0| 5.8558|?????/?| 4.2000 TO  6.4000|N||F||||20200407122232\r\u00034B\r\n\u0002" +
                    "0L|1|N\r\u000303\r\n\u0004";*/

                    Add_RTB(RTBout, $"{dtm} тест inputString:\n", Color.DarkGreen);
                    Add_RTB(RTBout, $"{inputString}", Color.Black);

                    st1_1 = ""; st2_1 = "";
                    CntParam = 0;
                    string pattern = ETX + "..";    // для удаления ETX и следующих за ним двух байт контрольной суммы - они не нужны мне.
                    inputString =Regex.Replace(inputString, pattern, String.Empty);

                    string[] s = inputString.Split(new string[] { CR, LF, ENQ, STX, EOT }, StringSplitOptions.RemoveEmptyEntries);
                    int ks = s.Count();
                    int i = 0;
                    string qq = "\x22"; // двойная кавычка
                    foreach (string si in s)
                    {
                        i++;
                        Add_RTB(RTBout, $"\n{i}-й: {si}", Color.Blue);

                    }
                    /*  результат s[i]:
                         3: 4R|1|^^^1^???????^0| 5.8558|?????/?| 4.2000 TO  6.4000|N||F||||20200407122232
                         4: 5R|2|^^^5^??.???^0|47.4037|?/?|64.0000 TO 83.0000|L||F||||20200407122031
                         5: 6R|3|^^^7^?????-?^0|50.3494|??????/?|71.0000 TO 115.0000|L||F||||20200407122116
                         6: 7R|4|^^^8^????????^0| 3.3998|?????/?| 1.7000 TO  8.3000|N||F||||20200407121945
                         7: 0R|5|^^^9^???^0|26.0191|??/?| 0.0000 TO 40.0000|N||F||||20200407122101
                         8: 1R|6|^^^10^???^0|57.2710|??/?| 0.0000 TO 37.0000|H||F||||20200407122046
                         9: 2R|7|^^^11^???????^0|15.5304|??/?|28.0000 TO 100.0000|L||F||||20200407122000
                        10: 3R|8|^^^12^???.???^0|180.3266|??/?|24.0000 TO 195.0000|N||F||||20200407122202
                        11: 4R|9|^^^13^???-??^0|46.7134|??/?| 0.0000 TO 25.0000|H||F||||20200407122217
                        12: 5R|10|^^^14^???^0|709.1733|??/?|230.0000 TO 460.0000|H||F||||20200407122147
                        13: 6R|11|^^^15^???.???.^0|488.5570|??/?|98.0000 TO 279.0000|H||F||||20200407122131
                        14: 7R|12|^^^16^???.??^0|11.7090|??????/?| 0.0000 TO 21.0000|N||F||||20200407122016
                             0  1  2-Kod         3-sRes  4-       5-sMes             6 7 8901 12 DT_An
                    */

                    for (i = 0; i<ks; i++)  // по количеству строк в переданных результатах пациента
                    {
                      string[] sTek = s[i].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);

                      if (s[i].Substring(1, 1) == "P")    // Patient
                        {
                            HistoryNumber = sTek[3];  // ||56174 ^??????? ?.?  ???|
                            HistoryNumber = Regex.Replace(HistoryNumber, "\\^.*", String.Empty);    // удалить всё начиная с "^"
                            //st0=s

                        }  
                      
                      if (s[i].Substring(1,1)=="R")   // Result
                        {
                            //Add_RTB(RTBout, $"\n{i}: {s[i]}", Color.Brown);
                            string sKod = sTek[2].Substring(3,2);
                            sKod = Regex.Replace(sKod, "\\^", String.Empty);    // удалить строку "^"
                            sResult = sTek[3].Trim();
                            sEdizm  = sTek[4].Trim();
                            sRefVal = sTek[5].Trim();
                            sRefVal = Regex.Replace(sRefVal, @"\s+", " ");  // заменяет несколько подряд идущих пробелов ординарными
                            if (sRefVal.Length > 16)
                                sRefVal = "длина больше 16";
                            sFlag   = sTek[6];
                            dt_Done = sTek[8];

                            CntParam++;
                            st1_1 += $", ParamName{CntParam}, ParamValue{CntParam}, ParamRef{CntParam}";    // + ParamMgr{CntParam}
                            st2_1 += $", '{sKod}', '{sResult}', '{sRefVal}'";
                            Add_RTB(RTBout, $"\n{i} HistoryNumber={HistoryNumber}, sKod={sKod}, sResult={sResult}, sEdizm={sEdizm}, sRefVal={sRefVal}, dt_Done={dt_Done}.", Color.Brown);
                        }

                    }

                    // формирование строки SQL Insert...
                    ResultText =  inputString;
                    //ResultDate= Get Date()
                    st2_0 = $") values ({Analyzer_Id}, host_name(), {HistoryNumber}, GetDate(), {CntParam}, '{ResultText}' ";    // для строки Insert...
                    st0 = st1_0 + st1_1 + st2_0 + st2_1 + ");";

                    //Add_RTB(RTBout, $" SQL:\n{st0}",Color.Red);

                    // ToSQL/...
                    WLog("SQL: " + st0);
                    ToSQL(st0);

                    HistoryNumber = ""; sResult = ""; sEdizm = ""; sRefVal = ""; sFlag = ""; dt_Done = "";
                    CntParam = 0; ResultDate = "";

                    Add_RTB(RTBout, $"\n{dtm} Конец теста inputString:\n", Color.DarkGreen);
                    Stat1.Text = "Конец теста.";
                    Stat1.ForeColor = Color.PowderBlue;
                    //MessageBox.Show($"{sMes}", $"Выполнен {CmbTest.SelectedItem}");
                    break;
                #endregion --- тест inputString
                case "Show Last SQL text":    // 2020-002-11
                    MessageBox.Show($"{scriptSql}\n *** Параметры:\n{scriptSqlPar}\n ***", "Тест последнего SQL:"
                        , MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    break;
                default:
                    MessageBox.Show($"Нет теста для:  {CmbTest.SelectedItem}\n");
                    break;
                    #endregion --- done cases or not allready used
            }
            #endregion --- Тесты по кнопке Выполнить
            // ---
        }

    }
}
