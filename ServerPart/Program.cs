using RiskManagementLibrary;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;


namespace ServerPart
{
    class Program
    {
        public static RiskManagementLib rs;
        static void Main(string[] args)
        {
            rs = new RiskManagementLib();
            Conect();
        }

        private static void Conect()
        {
            IPHostEntry ipHostInfo = Dns.Resolve(Dns.GetHostName());
            IPAddress ipAddress = ipHostInfo.AddressList[0];
            // створення об’єкту порт+ IP-адреса
            IPEndPoint localEndPoint = new IPEndPoint(IPAddress.Parse("127.0.0.1"), 12000);

            Socket listener = new Socket(AddressFamily.InterNetwork,
                                 SocketType.Stream,
                                 ProtocolType.Tcp);
            listener.Bind(localEndPoint);
            while (true)
            {
                listener.Listen(100);
                Socket handler = listener.Accept();
                byte[] bytes = new byte[100000];
                int count;
               // String data = " ";
                    // отримання та надсилання даних
                    count = handler.Receive(bytes);
                    rs = (RiskManagementLib)Convertor.ByteArrayToObject(bytes, count);
                if(!Check())
                {
                    rs.isCorrect = 0;
                }
                Console.WriteLine("До БД доступається " + rs.Login);
                Byte[] SendBytes = Convertor.ObjectToByteArray(rs);
                handler.Send(SendBytes);
                handler.Shutdown(SocketShutdown.Both);
                handler.Close();
            }
        }

        private static bool Check()
        {
            if(DB.SelectUsers(rs))
            {
                if(rs.isCorrect == 1)
                {
                    SelectData();
                }
                else
                {
                    Insert();
                }
                return true;
            }
            return false;
        }

        private static void SelectData()
        {
            DB.SelectIdentifyRisk(rs);
            DB.SelectEventIdentifyRisk(rs);
            DB.SelectAnalisIdentifyRisk(rs);
            DB.SelectAnalisRiskWithValidity(rs);
        }

        private static void Insert()
        {
            DB.InsertToIdentifyRisk(rs);
            DB.InsertToEventIdentifyRisk(rs);
            DB.InsertToAnalisIdentifyRisk(rs);
            DB.InsertToAnalisRiskWithValidity(rs);

        }
    }
}
