using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;


namespace AISpeakerServer
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Net.Sockets.UdpClient udp = null;
            Console.WriteLine("プログラムスタート");

            //バインドするローカルIPとポート番号
            try {
                string localIpString = "192.168.11.53";
                System.Net.IPAddress localAddress =
                    System.Net.IPAddress.Parse(localIpString);
                int localPort = 7000;

                //UdpClientを作成し、ローカルエンドポイントにバインドする
                Console.WriteLine("IP Address = " + localIpString + " Port = " + localPort + " UDPで待っています");
                Console.WriteLine();
                System.Net.IPEndPoint localEP =
                    new System.Net.IPEndPoint(localAddress, localPort);
                udp = new System.Net.Sockets.UdpClient(localEP);
            } catch (Exception)
            {
                Console.WriteLine("ネットワークのエラー");
                Console.WriteLine("何かキーを押してください");
                Console.ReadKey();
                Environment.Exit(0);
                
            }

            Word.Application word = null;
            Document document = null;
            object filename = System.IO.Directory.GetCurrentDirectory() + @"\out.docx";
            for (; ; )
            {
                //データを受信する
                System.Net.IPEndPoint remoteEP = null;
                byte[] rcvBytes = udp.Receive(ref remoteEP);

                //データを文字列に変換する
                string rcvMsg = System.Text.Encoding.UTF8.GetString(rcvBytes);

                //受信したデータと送信者の情報を表示する
                Console.WriteLine("受信したデータ:{0}", rcvMsg);
                Console.WriteLine("送信元アドレス:{0}/ポート番号:{1}",
                    remoteEP.Address, remoteEP.Port);


                //"open word"
                if (rcvMsg.Equals("open word\n") && word != null)
                {
                    // Word アプリケーションオブジェクトを作成
                    word = new Word.Application();

                    word.Visible = true;

                    // 新規文書を作成
                    document = word.Documents.Open(filename);

                    continue;
                }


                //"close word"
                if ((rcvMsg.Equals("close word\n") && word != null))
                {
                    //object filename = System.IO.Directory.GetCurrentDirectory() + @"\out.docx";
                    document.SaveAs2(ref filename);

                    // 文書を閉じる
                    document.Close();
                    document = null;
                    word.Quit();
                    word = null;

                    continue;
                }


                //"exit"を受信したら終了
                if (rcvMsg.Equals("exit\n"))
                {
                    Console.WriteLine("終了コマンド");
                    break;
                }


                if (word != null)
                {
                    addTextSample(ref document, WdColorIndex.wdBlack, rcvMsg);
                }

            }

            //UdpClientを閉じる
            udp.Close();

            Console.WriteLine("終了しました。");
            Console.ReadLine();
            }


        private static int getLastPosition(ref Document document)
        {
            return document.Content.End - 1;
        }
        private static void addTextSample(ref Document document, WdColorIndex color, string text)
        {
            int before = getLastPosition(ref document);
            Range rng = document.Range(document.Content.End - 1, document.Content.End - 1);
            rng.Text += text;
            int after = getLastPosition(ref document);

            document.Range(before, after).Font.ColorIndex = color;
        }
    }
}
