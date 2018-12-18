#define LOG_OUT // ログ出力

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
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
        string data = "";
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
        /* 機能     : フォーム表示時のイベント		                            */
        /* 引数     : なし                                                      */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void Form1_Shown(object sender, EventArgs e)
        {
            log_output("[START]SetRealTimeClockApp");

            serialPort.PortName = "COM5";   // 選択されたCOMをポート名に設定
            serialPort.Open();  // ポートを開く

            Task.Run(() => writeData());    // 別タスクで時刻設定コマンド送信
        }

        /************************************************************************/
        /* 関数名   : FormMain_FormClosing             			    			*/
        /* 機能     : フォームを閉じる時のイベント	                            */
        /* 引数     : なし                                                      */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            serialPort.Close();
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

            System.Threading.Thread.Sleep(10000);

            //時刻設定コマンド送信(リトライ10回)
            do
            {
                log_output("writeData:" + i);
                serialPort.Write(param, 0, param.Length);
                System.Threading.Thread.Sleep(3000);
                i++;
            } while (!ret && i < 10);
            
            endProcessing();
        }

        /************************************************************************/
        /* 関数名   : setClock_button_Click         						    */
        /* 機能     : ボタンクリック時のイベント	                            */
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
            int beginIndex;
            int endIndex;

            try
            {
                data += serialPort.ReadExisting();	// ポートから文字列を受信する
                log_output(data);
                if(!ret)
                { //システム時刻未設定
                    if (0 <= data.IndexOf("Y") && 0 <= data.IndexOf("END"))
                    { // 受信データに"Y"と"END"が含まれている
                        // 年
                        beginIndex = data.IndexOf("Y");
                        endIndex = data.IndexOf("M");
                        log_output(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        sysTime.wYear = Convert.ToUInt16(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        // 月
                        beginIndex = endIndex;
                        endIndex = data.IndexOf("D");
                        log_output(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        sysTime.wMonth = Convert.ToUInt16(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        // 日
                        beginIndex = endIndex;
                        endIndex = data.IndexOf("h");
                        log_output(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        sysTime.wDay = Convert.ToUInt16(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        // 時
                        beginIndex = endIndex;
                        endIndex = data.IndexOf("m");
                        log_output(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        sysTime.wHour = Convert.ToUInt16(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        // 分
                        beginIndex = endIndex;
                        endIndex = data.IndexOf("s");
                        log_output(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        sysTime.wMinute = Convert.ToUInt16(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        // 秒
                        beginIndex = endIndex;
                        endIndex = data.IndexOf("END");
                        log_output(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        sysTime.wSecond = Convert.ToUInt16(data.Substring((beginIndex + 1), (endIndex - 1 - beginIndex)));
                        sysTime.wMiliseconds = 0;
                        ret = SetLocalTime(ref sysTime);
                        log_output("[RTC]Set : " + ret);
                    }
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
            int i = 0;

            log_output("endProcessing");
            serialPort.Close();

            // SleepCheckerAppが起動していなかったら起動させる(リトライ5回)
            while (Process.GetProcessesByName("SleepCheckerApp").Length <= 0 && i < 5)
            {
                log_output("[ProcessStart]SleepCheckerApp");
                System.Diagnostics.Process p = System.Diagnostics.Process.Start("C:\\PCAnalysisApp\\SleepCheckerApp.exe");
                System.Threading.Thread.Sleep(300);
                i++;
            }

            if(i >= 5)
            { // リトライ5回した場合はWindows再起動
                log_output("[Reboot]");
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = "shutdown.exe";
                psi.Arguments = "-r -t 0";   // reboot
                psi.CreateNoWindow = true;
                Process p = Process.Start(psi);
            }
            log_output("[CLOSE]SetRealTimeClockApp");
            this.Close();
        }

        /************************************************************************/
        /* 関数名   : log_output          			    		    			*/
        /* 機能     : ログ出力                                                  */
        /* 引数     : [string] msg - ログ文言                                   */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void log_output(string msg)
        {
#if LOG_OUT
            Logging(logPath, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + "    " + msg + Environment.NewLine);
#endif
        }

        /************************************************************************/
        /* 関数名   : Logging          			    		          			*/
        /* 機能     : ログ書き込み処理                                          */
        /* 引数     : [string] logFullPath - ログ出力先パス                     */
        /*          : [string] logstr - ログ文言                                */
        /* 戻り値   : なし														*/
        /************************************************************************/
        private void Logging(string logFullPath, string logstr)
        {
            FileStream fs = new FileStream(logFullPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
            StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("Shift_JIS"));
            TextWriter tw = TextWriter.Synchronized(sw);
            tw.Write(logstr);
            tw.Flush();
            fs.Close();
        }
    }
}
