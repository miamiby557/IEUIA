﻿using mshtml;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
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
                    dictionary.Add("x", offsetLeft);
                    dictionary.Add("y", offsetTop);
                    dictionary.Add("w", offsetWidth);
                    dictionary.Add("h", offsetHeight);

                    // 截图
                    Console.WriteLine("SaveImage:" + offsetLeft.ToString() + ",offsetTop:" + offsetTop.ToString() + ",offsetWidth:" + offsetWidth.ToString() + ",offsetHeight:" + offsetHeight.ToString());
                    // Bitmap image = this.SaveImage(offsetLeft - 100, offsetTop - 50, offsetWidth + 50,offsetHeight+ 50
                    MouseHookHelper.SetCursorPos(0, 0);
                    Thread.Sleep(150);
                    String imagePath = this.SaveImage(offsetLeft, offsetTop, offsetWidth, offsetHeight);
                    // string base64FromImage = ImageUtil.GetBase64FromImage(image);
                    dictionary.Add("screenShot", imagePath);

                    PostData(dictionary);
                }
            }
        }

        // 保存截图
        private String SaveImage(int x, int y, int width, int height)
        {
            Bitmap bitmap = new Bitmap(width, height);
            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
            String imagePath = Directory.GetCurrentDirectory() + "\\img.png";
            bitmap.Save(imagePath, ImageFormat.Png);
            return imagePath;
        }


        private Object o = new Object();

        private void PostData(Dictionary<string, object> dictionary)
        {
            try
            {
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