using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;


namespace SubtitleThrower
{
    public partial class Form1 : Form
    {
        // global variables
        public int prevstartread, startread, prevendread, endread;
        public static string stopset = "，。？！」”、：；.,:;?!]\"";
        public string processname = "aegisub32";  // set default value for processname
        public string richtextboxtext;
        public string file_sel_pos = "AutoTextHighlighterSelectionStart.txt",
                      file_text = "AutoTextHighlighterSave.txt";

        const int MYACTION_HOTKEY_F4 = 4;
        const int MYACTION_HOTKEY_F5 = 5;
        const int MYACTION_HOTKEY_F6 = 6;
        const int MYACTION_HOTKEY_F7 = 7;

        // DLL libraries used to manage hotkeys
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);



        // FindWindowEx & SendMessage method. Not being used for now. Maybe in the future
        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, System.Text.StringBuilder text);

        // Use SetForegroundWindow & SendKeys instead of FindWindowEx & SendMessage
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);


        public Form1()
        {
            InitializeComponent();

            // Get all current running processes into Combo Box1
            comboBox1.Items.Clear();
            Process[] MyProcess = Process.GetProcesses();
            for (int i = 0; i < MyProcess.Length; i++)
                comboBox1.Items.Add(MyProcess[i].ProcessName);
            // comboBox1.Items.Add(MyProcess[i].ProcessName + "-" + MyProcess[i].Id);
            comboBox1.Sorted = true;


            // Modifier keys codes: Alt = 1, Ctrl = 2, Shift = 4, Win = 8
            // Compute the addition of each combination of the keys you want to be pressed
            // ALT+CTRL = 1 + 2 = 3 , CTRL+SHIFT = 2 + 4 = 6...
            RegisterHotKey(this.Handle, MYACTION_HOTKEY_F4, 0, (int)Keys.F4);
            RegisterHotKey(this.Handle, MYACTION_HOTKEY_F5, 0, (int)Keys.F5);
            RegisterHotKey(this.Handle, MYACTION_HOTKEY_F6, 0, (int)Keys.F6);
            RegisterHotKey(this.Handle, MYACTION_HOTKEY_F7, 0, (int)Keys.F7);


            // Handle the ApplicationExit event to know when the application is exiting.
            Application.ApplicationExit += new EventHandler(this.OnApplicationExit);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == MYACTION_HOTKEY_F4)
            {
                // My hotkey has been typed
                button1.PerformClick();
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == MYACTION_HOTKEY_F5)
            {
                // My hotkey has been typed
                button5.PerformClick();
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == MYACTION_HOTKEY_F6)
            {
                // My hotkey has been typed
                button4.PerformClick();
            }
            else if (m.Msg == 0x0312 && m.WParam.ToInt32() == MYACTION_HOTKEY_F7)
            {
                // My hotkey has been typed
                button7.PerformClick();
            }
            base.WndProc(ref m);
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            richTextBox1.Text = System.IO.File.ReadAllText(@file_text);
            richTextBox1.SelectionStart = int.Parse(System.IO.File.ReadAllText(@file_sel_pos));

            startread = richTextBox1.SelectionStart;

            richTextBox1.ScrollToCaret();
            
            // Progress Bar
            progressBar1.Value = (int)Math.Round(Convert.ToDecimal(100*richTextBox1.SelectionStart/richTextBox1.TextLength), MidpointRounding.AwayFromZero);
            label2.Text = richTextBox1.SelectionStart.ToString() +　"/" + richTextBox1.TextLength.ToString();
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            // When the application is exiting...
            System.IO.File.WriteAllText(@file_sel_pos, startread.ToString());
            System.IO.File.WriteAllText(@file_text, richtextboxtext);
            // UnregisterHotKey(this.Handle, MYACTION_HOTKEY_F4);
            // UnregisterHotKey(this.Handle, MYACTION_HOTKEY_F6);
        }


        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richtextboxtext = richTextBox1.Text;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // COPY OR EXPERIMENTAL BUTTOM
            //Clipboard.SetText(richTextBox1.SelectedText);

            // SEND TEXT Button
            // Not being used for now. Maybe in the future
            // FindWindowEx & SendMessage method.
            /*
            Process[] notepads = Process.GetProcessesByName("aegisub32");
            if (notepads.Length == 0)
            {
                MessageBox.Show("windows not found");
                return;
            }
            if (notepads[0] != null)
            {
                IntPtr child = FindWindowEx(notepads[0].MainWindowHandle, new IntPtr(0), "stcwindow", null);
                System.Text.StringBuilder sb = new System.Text.StringBuilder(255);
                SendMessage(new IntPtr(0x00E812E8), 0x0D, 11, sb);
                // SendMessage(new IntPtr(0x00301258), 0x0C, 0, sb);

                // SEND {ENTER}
                // Use SetForegroundWindow & SendKeys instead of FindWindowEx & SendMessage
                if (processname == null)
                {
                    MessageBox.Show("No Target Window Selected");
                    return;
                }
                Process p = Process.GetProcessesByName(processname).FirstOrDefault();
                if (p != null)
                {
                    // send {ENTER}
                    IntPtr h = p.MainWindowHandle;
                    
                    SetForegroundWindow(h);
                    SendKeys.SendWait("^4");
                    // SendKeys.SendWait(sb.ToString());
                }

            }*/
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Process p = Process.GetProcessesByName(processname).FirstOrDefault();
            if (p != null)
            {
                // send {ENTER}
                IntPtr h = p.MainWindowHandle;

                SetForegroundWindow(h);
                SendKeys.SendWait("^4");
                SendKeys.SendWait("{ENTER}");
                SendKeys.SendWait("^p");
                if (checkBox2.Checked)  // Send ctrl-s save command
                {
                    SendKeys.SendWait("^s");
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            startread = richTextBox1.SelectionStart;
            System.IO.File.WriteAllText(@file_sel_pos, startread.ToString());
            System.IO.File.WriteAllText(@file_text, richtextboxtext);
            MessageBox.Show("Content and Position saved! ");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            // SEND CTRL-S Save command
            // Use SetForegroundWindow & SendKeys instead of FindWindowEx & SendMessage
            if (processname == null)
            {
                MessageBox.Show("No Target Window Selected");
                return;
            }
            Process p = Process.GetProcessesByName(processname).FirstOrDefault();
            if (p != null)
            {
                // send {ENTER}
                IntPtr h = p.MainWindowHandle;
                SetForegroundWindow(h);
                SendKeys.SendWait("^s");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            processname = comboBox1.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // SEND {ENTER}
            // Use SetForegroundWindow & SendKeys instead of FindWindowEx & SendMessage
            if (processname == null)
            {
                MessageBox.Show("No Target Window Selected");
                return;
            }
            Process p = Process.GetProcessesByName(processname).FirstOrDefault();
            if (p != null)
            {
                // send {ENTER}
                IntPtr h = p.MainWindowHandle;
                SetForegroundWindow(h);
                SendKeys.SendWait("{Enter}");
            }


                // SEND TEXT Button
                // Not being used for now. Maybe in the future
                // FindWindowEx & SendMessage method.
                /*
                Process[] notepads = Process.GetProcessesByName("aegisub32");
                if (notepads.Length == 0)
                {
                    MessageBox.Show("windows not found");
                    return;
                }
                if (notepads[0] != null)
                {
                    IntPtr child = FindWindowEx(notepads[0].MainWindowHandle, new IntPtr(0), "stcwindow", null);
                    SendMessage(child, 0x000C, 0, richTextBox1.SelectedText);

                }*/

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Previous Button
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Next Button

            // check if cursor already reached the end.
            if (richTextBox1.SelectionStart == richTextBox1.TextLength)
            {
                MessageBox.Show("Done! End of text reached!!!");
                return;
            }

            // Use SetForegroundWindow & SendKeys instead of FindWindowEx & SendMessage
            if (processname == null)
            {
                MessageBox.Show("No Target Window Selected");
                return;
            }
            Process p = Process.GetProcessesByName(processname).FirstOrDefault();
            if (p != null)
            {
                // get selected text
                startread = richTextBox1.SelectionStart;
                endread = startread;

                while (endread < richTextBox1.TextLength &&
                    !stopset.Contains(richTextBox1.Text[endread].ToString()))
                {
                    endread++;
                }
                while (endread+1 < richTextBox1.TextLength &&
                     stopset.Contains(richTextBox1.Text[endread+1].ToString()))
                {
                    endread++;
                }
                // remove previous highlight
                richTextBox1.Select(prevstartread, prevendread - prevstartread + 1);
                richTextBox1.SelectionBackColor = Color.White;
                richTextBox1.SelectionColor = Color.Black;

                // highlight selected text
                richTextBox1.Select(startread, endread - startread + 1);
                richTextBox1.SelectionBackColor = Color.Blue;
                richTextBox1.SelectionColor = Color.White;
                prevstartread = startread;
                prevendread = endread;

                // copy text to clipboard
                Clipboard.SetText(richTextBox1.SelectedText);

                // send keys
                IntPtr h = p.MainWindowHandle;
                SetForegroundWindow(h);
                SendKeys.SendWait(richTextBox1.SelectedText);
                if (checkBox1.Checked)  // Send ENTER
                {
                    SendKeys.SendWait("{Enter}");
                }
                richTextBox1.ScrollToCaret();
                endread++;
                richTextBox1.SelectionStart = endread;
                while (endread < richTextBox1.TextLength &&
                    richTextBox1.Text[richTextBox1.SelectionStart].ToString() == "\n")
                {
                    richTextBox1.SelectionStart = richTextBox1.SelectionStart + 1;
                }
                startread = richTextBox1.SelectionStart;
                progressBar1.Value = (int)Math.Round(Convert.ToDecimal(100 * richTextBox1.SelectionStart / richTextBox1.TextLength), MidpointRounding.AwayFromZero);
                label2.Text = richTextBox1.SelectionStart.ToString() + "/" + richTextBox1.TextLength.ToString();
            }
        }
    }
}
