using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TimeSync
{
    class Program
    {
        const string TimeFormat = "O";
        static void Main(string[] args)
        {
            if (!IsRunAsAdmin())
            {
                Console.WriteLine("請使用管理員身分執行。");
            }
            try
            {
                if (args.Contains("w32time"))
                {
                    Console.WriteLine($"訊息：透過 win32time service 更新時間");
                    SyncDateTime();
                    return;
                }
                var time = args.SingleOrDefault(arg => arg.StartsWith("time:"));
                if (!string.IsNullOrEmpty(time))
                {
                    time = time.Replace("time:", "");
                    Console.WriteLine($"訊息：透過指定時間戳記進行設定 {time}");
                    long fileTime = Convert.ToInt64(time);
                    Console.WriteLine($"time:{DateTime.FromFileTime(fileTime)}");
                    DateTime dateTime = DateTime.FromFileTime(fileTime);
                    SetLocalDateTime(dateTime);
                    return;
                }
                var ntpServer = args.SingleOrDefault(arg => arg.StartsWith("ntp:"));
                if (!string.IsNullOrEmpty(ntpServer))
                {
                    ntpServer = ntpServer.Replace("ntp:", "");
                    Console.WriteLine($"訊息：透過指定NTP伺服器 {ntpServer} 進行設定");
                }
                else
                {
                    ntpServer = "time.google.com";
                    Console.WriteLine($"訊息：透過預設NTP伺服器 {ntpServer} 進行設定");
                }
                SetLocalDateTime(GetNetworkTime(ntpServer));
            }
            catch(Exception ex)
            {
                Console.WriteLine($"訊息：發生錯誤:{ex.Message}");
            }
            finally
            {
                Thread.Sleep(2500);
            }
        }
        private static bool IsRunAsAdmin()
        {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(id);

            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        [DllImport("kernel32.dll")]
        private static extern bool SetLocalTime(ref Systemtime time);

        [StructLayout(LayoutKind.Sequential)]
        private struct Systemtime
        {
            public short year;
            public short month;
            public short dayOfWeek;
            public short day;
            public short hour;
            public short minute;
            public short second;
            public short milliseconds;
        }

        /// <summary>
        /// 設定系統時間
        /// </summary>
        /// <param name="dt">需要設定的時間</param>
        /// <returns>返回系統時間設定狀態，true為成功，false為失敗</returns>
        public static bool SetLocalDateTime(DateTime dt)
        {
            Systemtime st;

            st.year = (short)dt.Year;
            st.month = (short)dt.Month;
            st.dayOfWeek = (short)dt.DayOfWeek;
            st.day = (short)dt.Day;
            st.hour = (short)dt.Hour;
            st.minute = (short)dt.Minute;
            st.second = (short)dt.Second;
            st.milliseconds = (short)dt.Millisecond;
            bool rt = SetLocalTime(ref st);
            if (rt)
            {
                Console.WriteLine($"訊息：已更新系統時間為：{dt.ToLongDateString()} {dt.ToLocalTime()}");
            }
            else
            {
                Console.WriteLine($"訊息：更新系統時間失敗，欲設定的時間為：{dt.ToLongDateString()} {dt.ToLocalTime()}");
            }
            return rt;
        }
        public static DateTime GetNetworkTime(string ntpServer)
        {
            // NTP message size - 16 bytes of the digest (RFC 2030)
            var ntpData = new byte[48];

            //Setting the Leap Indicator, Version Number and Mode values
            ntpData[0] = 0x1B; //LI = 0 (no warning), VN = 3 (IPv4 only), Mode = 3 (Client Mode)

            var addresses = Dns.GetHostEntry(ntpServer).AddressList;

            //The UDP port number assigned to NTP is 123
            var ipEndPoint = new IPEndPoint(addresses[0], 123);
            //NTP uses UDP

            using (var socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp))
            {
                socket.Connect(ipEndPoint);

                //Stops code hang if NTP is blocked
                socket.ReceiveTimeout = 3000;

                socket.Send(ntpData);
                socket.Receive(ntpData);
                socket.Close();
            }

            //Offset to get to the "Transmit Timestamp" field (time at which the reply 
            //departed the server for the client, in 64-bit timestamp format."
            const byte serverReplyTime = 40;

            //Get the seconds part
            ulong intPart = BitConverter.ToUInt32(ntpData, serverReplyTime);

            //Get the seconds fraction
            ulong fractPart = BitConverter.ToUInt32(ntpData, serverReplyTime + 4);

            //Convert From big-endian to little-endian
            intPart = SwapEndianness(intPart);
            fractPart = SwapEndianness(fractPart);

            var milliseconds = (intPart * 1000) + ((fractPart * 1000) / 0x100000000L);

            //**UTC** time
            var networkDateTime = (new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Utc)).AddMilliseconds((long)milliseconds);

            return networkDateTime.ToLocalTime();
        }

        // stackoverflow.com/a/3294698/162671
        static uint SwapEndianness(ulong x)
        {
            return (uint)(((x & 0x000000ff) << 24) +
                           ((x & 0x0000ff00) << 8) +
                           ((x & 0x00ff0000) >> 8) +
                           ((x & 0xff000000) >> 24));
        }
        public static bool SyncDateTime()
        {
            try
            {
                ServiceController serviceController = new ServiceController("w32time");

                if (serviceController.Status != ServiceControllerStatus.Running)
                {
                    serviceController.Start();
                }

                Console.WriteLine("訊息：w32time service is running");

                Process processTime = new Process();
                processTime.StartInfo.FileName = "w32tm";
                processTime.StartInfo.Arguments = "/resync";
                processTime.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                processTime.Start();
                processTime.WaitForExit();

                Console.WriteLine("訊息：w32time service has sync local dateTime from NTP server");

                return true;
            }
            catch (Exception exception)
            {
                Console.WriteLine("訊息：unable to sync date time from NTP server", exception);

                return false;
            }
        }
    }
}
