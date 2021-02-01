using mshtml;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace IEXPath
{
    partial class frmIEXPath
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

 #region Windows 窗体设计器生成的代码

        private KeyEventHandler myKeyEventHandeler = null;//按键钩子
        private KeyboardHook k_hook = new KeyboardHook();


        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.Text = "IE浏览器拾取插件";
            base.Name = "CINDATA AUTOMATION";
            base.ClientSize = new System.Drawing.Size(200, 30);
            base.StartPosition = FormStartPosition.Manual;
            Rectangle rect = Screen.GetWorkingArea(this);
            System.Drawing.Point p = new System.Drawing.Point(rect.Width - 250, rect.Height - 40);
            this.Location = p;
            this.TopMost = true;
            this.ControlBox = false;

            this.timer1.Interval = 300;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            timer1.Enabled = true;
            this.Cursor = this.moveCursor;

            myKeyEventHandeler = new KeyEventHandler(hook_KeyDown);
            k_hook.KeyDownEvent += myKeyEventHandeler;//钩住键按下
            k_hook.Start();//安装键盘钩子

        }

        public static IHTMLElement element;

        private void hook_KeyDown(object sender, KeyEventArgs e)
        {
            //  这里写具体实现
            Console.WriteLine("按下按键" + e.KeyValue);
            if (e.KeyValue == 27)
            {
                // 关闭
                Thread.Sleep(400);
                this.timer1.Stop();
                clearBorderHint();
                this.timer1.Stop();
                this.Close();
                Application.Exit();
            }
            else if (e.KeyValue == 162)
            {
                lock (o)
                {
                    Dictionary<string, object> dictionary = new Dictionary<string, object>();
                    dictionary.Add("xpath", xpath);
                    dictionary.Add("txtID", txtID);
                    dictionary.Add("txtName", txtName);
                    dictionary.Add("tagName", tagName);
                    dictionary.Add("txtText", txtText);
                    dictionary.Add("txtClass", txtClass);
                    dictionary.Add("txtHTML", txtHTML);
                    dictionary.Add("txtOuterHtml", txtOuterHtml);
                    PostData(dictionary);
                }
            }
        }


        private Object o = new Object();

        private void PostData(Dictionary<string, object> dictionary)
        {
            try
            {
                /*string url = "http://localhost:63361/postData";
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "POST";
                req.Timeout = 4000;//设置请求超时时间，单位为毫秒
                req.ContentType = "application/json";
                byte[] data = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(dictionary));
                req.ContentLength = data.Length;
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(data, 0, data.Length);
                    reqStream.Close();
                    // 关闭
                    Thread.Sleep(400);
                    this.timer1.Stop();
                    clearBorderHint();
                    this.Close();
                    Application.Exit();
                }*/
                Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                //连接服务器
                socket.Connect(new IPEndPoint(IPAddress.Parse("127.0.0.1"), 10083));
                dictionary.Add("dataType", "IE_UIA");
                socket.Send(Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(dictionary)));
                socket.Close();
                // 关闭
                Thread.Sleep(400);
                this.timer1.Stop();
                clearBorderHint();
                this.Close();
                Application.Exit();
            }
            catch (Exception)
            {
                // 关闭
                Thread.Sleep(400);
                this.timer1.Stop();
                clearBorderHint();
                this.Close();
                Application.Exit();
            }
        }

        private System.Windows.Forms.ListBox lstXPath = new ListBox();
        private System.Windows.Forms.Timer timer1;
    }
}
#endregion