using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using Sharp7;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using static System.Net.Mime.MediaTypeNames;

namespace SprayingCabineService
{
    public class SprayingCabineService : BackgroundService
    {
        private readonly ILogger<SprayingCabineService> _logger;
        private static PeriodicTimer timer;
        private static S7Client PLCclient = new S7Client();
        private static Dictionary<string, string> ConfigSettings = new Dictionary<string, string>();
        private static byte[] Buffer = new byte[65536];

        private int processing = 0;
        private string barCode = string.Empty;
        private DataTable dataTable = new DataTable();
        private string m3code = string.Empty;
        private bool setOK = false;

        public SprayingCabineService(ILogger<SprayingCabineService> logger) { _logger = logger; }


        //Processing Cycle
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {

            //Load AND Check Setting
            LoadSettings();
            CheckSettings();

            timer = new PeriodicTimer(TimeSpan.FromMilliseconds(double.Parse(ConfigSettings.GetValueOrDefault("stepInterval"))));
            int ConnectionResult = PLCclient.ConnectTo(ConfigSettings.GetValueOrDefault("plcAddress"), int.Parse(ConfigSettings.GetValueOrDefault("rack")), int.Parse(ConfigSettings.GetValueOrDefault("slot")));
            int PlcResult = 1;
            int BarCodeResult = 1;

            try {
                //Proccess Cycle
                while (await timer.WaitForNextTickAsync() && true)
                {
                    //PLC Connection
                    if (processing == 0 && ConnectionResult == 0) {
                        //WriteLogFile("PLC Connected");
                        processing = 5;
                    } else if (processing == 0 && ConnectionResult != 0) {
                        processing = 0;
                        WriteLogFile("PLC Connection Error: " + ConnectionResult.ToString() + " ->" + PLCclient.ErrorText(ConnectionResult));
                        ConnectionResult = PLCclient.ConnectTo(ConfigSettings.GetValueOrDefault("plcAddress"), 0, 0);
                    }


                    //Check Process
                    if (processing == 5) {
                        //Load Start Request
                        Buffer = new byte[1];
                        PlcResult = PLCclient.DBRead(2800, 0, 1, Buffer);


                        /* Result.SerialNumber = S7.GetCharsAt(Buffer, 0, 12)  Result.TestResult = S7.GetIntAt(Buffer, 12) Result.LeakDetected = S7.GetRealAt(Buffer, 14) */

                        //Start Process
                        if (PlcResult == 0 && S7.GetBitAt(Buffer, 0, 0)) {
                            processing = 10;
                            WriteLogFile("Process Started: " + ByteArrayToString(Buffer));

                            //Load ScanCode
                            Buffer = new byte[int.Parse(ConfigSettings.GetValueOrDefault("barCodeLength"))];
                            BarCodeResult = PLCclient.DBRead(2800, int.Parse(ConfigSettings.GetValueOrDefault("barCodeStart")), int.Parse(ConfigSettings.GetValueOrDefault("barCodeLength")), Buffer);

                        } else if (PlcResult == 0 && !S7.GetBitAt(Buffer, 0, 0)) { processing = 0;
                        } else if (PlcResult != 0) { WriteLogFile("Query Requested PLC Error: " + PlcResult.ToString() + " ->" + PLCclient.ErrorText(PlcResult) + ": " + ByteArrayToString(Buffer)); }
                        
                    }


                    //Check ScanCode
                    if (processing == 10 && BarCodeResult == 0) {
                        processing = 20;
                        barCode = S7.GetCharsAt(Buffer, 0, int.Parse(ConfigSettings.GetValueOrDefault("barCodeLength")));
                        WriteLogFile("PLC BarCode Loaded: " + barCode);
                    } else if (processing == 10 && BarCodeResult != 0) { processing = 0; WriteLogFile("ScanCode Requested PLC Error: " + BarCodeResult.ToString() + " ->" + PLCclient.ErrorText(BarCodeResult) + ": " + ByteArrayToString(Buffer)); }


                    //Get Sql M3 Code
                    if (processing == 20) {
                        SqlConnection cnn = new SqlConnection(ConfigSettings.GetValueOrDefault("connectionString"));
                        cnn.Open();

                        try {
                                                       
                            if (cnn.State == ConnectionState.Open) {
                                SqlDataAdapter mDataAdapter = new SqlDataAdapter(new SqlCommand("SELECT TOP 1 a.[code] as 'M3 Code' FROM [PILM].[codebook].[product_type] a ,[PILM].[audit].[production__product] b WHERE b.product_type_id = a.id AND b.code = '" + barCode + "'", cnn));
                                dataTable.Clear();
                                mDataAdapter.Fill(dataTable);
                                m3code = dataTable.Rows[0].ItemArray[0].ToString();
                                cnn.Close();
                                setOK = true;
                                WriteLogFile("SQL M3code Query Found: " + m3code);
                                
                            } else { WriteLogFile("SQL M3code Query Error: Not connected"); setOK = false; }
                          
                        } catch (Exception Ex) { 
                            WriteLogFile("SQL Query Error: barcode: " + barCode + Environment.NewLine + Ex.StackTrace); 
                            if (cnn.State == ConnectionState.Open) { cnn.Close(); }
                            setOK = false; 
                        }
                        processing = 30;
                    }


                    //Write Data To PLC AND set Process Finished
                    if (processing == 30) {
                        try {

                            if (setOK) {

                                // Set Empty M3 code
                                Buffer = new byte[int.Parse(ConfigSettings.GetValueOrDefault("m3CodeLength"))];
                                PlcResult = PLCclient.DBWrite(2800, int.Parse(ConfigSettings.GetValueOrDefault("m3CodeStart")), int.Parse(ConfigSettings.GetValueOrDefault("m3CodeLength")), Buffer);

                                // Set M3 code
                                Buffer = new byte[Encoding.ASCII.GetBytes(m3code).Length];
                                Buffer = Encoding.ASCII.GetBytes(m3code);
                                PlcResult = PLCclient.DBWrite(2800, int.Parse(ConfigSettings.GetValueOrDefault("m3CodeStart")), Encoding.ASCII.GetBytes(m3code).Length, Buffer);

                                //Set OK
                                Buffer = new byte[1]; Buffer[0] = 0x02;
                                PlcResult = PLCclient.DBWrite(2800, 0, 1, Buffer);

                            } else {

                                // Set Empty M3 code
                                Buffer = new byte[int.Parse(ConfigSettings.GetValueOrDefault("m3CodeLength"))];
                                PlcResult = PLCclient.DBWrite(2800, int.Parse(ConfigSettings.GetValueOrDefault("m3CodeStart")), int.Parse(ConfigSettings.GetValueOrDefault("m3CodeLength")), Buffer);

                                // Set NOK
                                Buffer = new byte[1]; Buffer[0] = 0x04;
                                PlcResult = PLCclient.DBWrite(2800, 0, 1, Buffer);
                            }

                            //Process Finished 
                            //Buffer = new byte[1]; Buffer[0] = 0x00;
                            //PlcResult = PLCclient.DBWrite(2800, 0, 1, Buffer);
                            
                        }
                        catch (Exception Ex) { WriteLogFile("PLC Write Error: Buffer: " + Buffer + Environment.NewLine + "message: "+ Ex.StackTrace); }
                        processing = 0;
                    }



                }








            } catch(Exception ex) { WriteLogFile("Program Exception: " + ex.StackTrace); }

            Debug.WriteLine(DateTime.Now);
            _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
        }




        //Load Config
        public static void LoadSettings() {
            try {
                string json = File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Data", "config.json"), GlobalFunctions.FileDetectEncoding(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Data", "config.json")));
                Dictionary<string, string> SettingData = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                ConfigSettings.Add(SettingData.Keys.ToList()[0].ToString(), SettingData.Values.ToList()[0].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[1].ToString(), SettingData.Values.ToList()[1].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[2].ToString(), SettingData.Values.ToList()[2].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[3].ToString(), SettingData.Values.ToList()[3].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[4].ToString(), SettingData.Values.ToList()[4].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[5].ToString(), SettingData.Values.ToList()[5].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[6].ToString(), SettingData.Values.ToList()[6].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[7].ToString(), SettingData.Values.ToList()[7].ToString());
                ConfigSettings.Add(SettingData.Keys.ToList()[8].ToString(), SettingData.Values.ToList()[8].ToString());
            } catch (Exception ex) { WriteLogFile("Data\\config.json Error"); }
        }


        //Check PLC and SQL
        public static void CheckSettings() {
            bool error = false;

            //Check PLC connection
            int plcConnResult = PLCclient.ConnectTo(ConfigSettings.GetValueOrDefault("plcAddress"), int.Parse(ConfigSettings.GetValueOrDefault("rack")), int.Parse(ConfigSettings.GetValueOrDefault("slot")));
            if (plcConnResult != 0) { WriteLogFile("PLC Connection Error Exiting");error = true; }
            else { WriteLogFile("PLC Connection is OK"); }

                //Check SQL connection
                try
                {
                    SqlConnection cnn = new SqlConnection(ConfigSettings.GetValueOrDefault("connectionString"));
                    cnn.Open();
                    if (cnn.State == ConnectionState.Open)
                    {
                        cnn.Close(); WriteLogFile("SQL Connection is OK");
                    }
                    else { WriteLogFile("SQL Connection Not Open Exiting"); error = true; }
                }
                catch (Exception ex) { WriteLogFile("SQL Connection Error Exiting: " + ex.StackTrace); error = true; }

            //Bad Configuration EXIT
            if (error) { Environment.Exit(0); }
        }


        //Write Log File
        public static void WriteLogFile(string message) {
            string log = string.Empty;
            try {
                log = File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Data", "datalog.txt"));
            } catch (Exception ex) { }
            try {
                File.WriteAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Data", "datalog.txt"), DateTimeOffset.Now.DateTime.ToUniversalTime().ToString() + ": " + message + Environment.NewLine + (log.Length > 1000000 ? log.Substring(0, 1000000) : log));
            } catch (Exception ex) { }
        }


        public static string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2}", b);
            return hex.ToString();
        }
    }
}
