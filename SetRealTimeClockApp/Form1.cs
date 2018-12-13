#define LOG_OUT // ログ出力

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SetRealTimeClockApp
{
    public struct SystemTime
    {
        public ushort wYear;
        public ushort wMonth;
        public ushort wDayOfWeek;
        public ushort wDay;
        public ushort wHour;
        public ushort wMinute;
        public ushort wSecond;
        public ushort wMiliseconds;
    }

    public partial class Form1 : Form
    {
        [DllImport("kernel32.dll")]
        public static extern bool SetLocalTime(ref SystemTime sysTime);

        SystemTime sysTime = new SystemTime();
        string clock = "";
//        string temp = "";
        bool ret = false;
        
        // ログ出力パス
        string logPath = "C:\\log";

        public Form1()
        {
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }
            logPath = logPath + "\\log.txt";

            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        /************************************************************************/
        /* 関数名   : Form1_Shown          						    			*/
        /* 機能     : フォーム表示時イベント		                            */
        /* 引数     : なし                                                      */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void Form1_Shown(object sender, EventArgs e)
        {
            serialPort.PortName = "COM5";   // 選択されたCOMをポート名に設定
            serialPort.Open();  // ポートを開く
            
            Task.Run(() => writeData());    // 別タスクで時刻設定コマンド送信
        }

        /************************************************************************/
        /* 関数名   : writeData          						    			*/
        /* 機能     : 時刻設定コマンド送信  		                            */
        /* 引数     : なし                                                      */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void writeData()
        {
            int i = 0;
            byte[] param = new byte[1];
            param[0] = (byte)2; // 時刻設定コード

            //時刻設定コマンド送信(リトライ5回)
            do
            {
                log_output("writeData:" + i);
                serialPort.Write(param, 0, param.Length);
                Task.Delay(1000);
                i++;
            } while (!ret && i < 5);
            
            endProcessing();
        }

        /************************************************************************/
        /* 関数名   : setClock_button_Click         						    */
        /* 機能     : ボタンクリック時イベント		                            */
        /* 引数     : なし                                                      */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void setClock_button_Click(object sender, EventArgs e)
        {
        }

        /************************************************************************/
        /* 関数名   : serialPort_DataReceived    				    			*/
        /* 機能     : Arduinoからのデータ受信関数	                            */
        /* 引数     : なし                                                      */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void serialPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            try
            {
                string data = serialPort.ReadExisting();	// ポートから文字列を受信する
                if (!string.IsNullOrEmpty(data))
                {
                    Invoke((MethodInvoker)(() =>	// 受信用スレッドから切り替えてデータを書き込む
                    {
                        switch(data)
                        {
                            case "YEAR":
                                sysTime.wYear = Convert.ToUInt16(clock);
                                break;
                            case "MONTH":
                                sysTime.wMonth = Convert.ToUInt16(clock);
                                break;
                            case "DAY":
                                sysTime.wDay = Convert.ToUInt16(clock);
                                break;
                            case "HOUR":
                                sysTime.wHour = Convert.ToUInt16(clock);
                                break;
                            case "MINUTE":
                                sysTime.wMinute = Convert.ToUInt16(clock);
                                break;
                            case "SECOND":
                                sysTime.wSecond = Convert.ToUInt16(clock);
                                sysTime.wMiliseconds = 0;
                                break;
                            case "END":
                                // システム時刻を設定する
                                ret = SetLocalTime(ref sysTime);
                                log_output("[RTC]Set : " + ret);
                                //                                label.Text = temp;
                                break;
                            default:
                                clock = data;
//                                temp = temp + data;
                                break;
                        }
                        Thread.Sleep(10);
                    }));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /************************************************************************/
        /* 関数名   : endProcessing          					    			*/
        /* 機能     : 終了処理                		                            */
        /* 引数     : なし                                                      */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void endProcessing()
        {
            log_output("endProcessing");
            serialPort.Close();
            System.Diagnostics.Process p = System.Diagnostics.Process.Start("C:\\PCAnalysisApp\\SleepCheckerApp.exe");
            log_output("[CLOSE]SetRealTimeClockApp");
            this.Close();
        }

        /************************************************************************/
        /* 関数名   : log_output          			    		    			*/
        /* 機能     : C:\\log\\log.txtに出力                                    */
        /* 引数     : string : msg - ログ文言                                   */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void log_output(string msg)
        {
#if LOG_OUT
            File.AppendAllText(logPath, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "    " + msg + Environment.NewLine);
#endif
        }
    }
}
