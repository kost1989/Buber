using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.IO.Compression;
using System.Drawing.Printing;
using System.Printing;
using System.Diagnostics;
using Ionic.Zip;



namespace buber
{
    public partial class Form1 : Form
    {
        int indx_tabl = 0;
        int threads_text = 0;
        int threads_left = 0;
        int p_sec = 0;
        int a_sec = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void start_copy()
        {
            ++threads_text;
            this.linkLabel1.BeginInvoke((MethodInvoker)(() => this.linkLabel1.Text = threads_text.ToString()));
            this.checkBox6.BeginInvoke((MethodInvoker)(() => checkBox6.Checked = true)); 
            //timer1.Enabled = false;
            bool flag = false;
            this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.Text = "Бубер - Работаем"));
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(textBox1.Text);
                DirectoryInfo dirInfo2 = new DirectoryInfo(textBox2.Text);
                DirectoryInfo dirTempInfo = new DirectoryInfo(@"C:\Temp\Buber");

                foreach (FileInfo fileDel in dirTempInfo.GetFiles())
                {
                    fileDel.Delete();
                }

                foreach (FileInfo file in dirInfo.GetFiles("*" + textBox3.Text + "*.*"))
                {
                    // РАСПАКОВЫВАЕМ И ПЕЧАТАЕМ И УДАЛЯЕМ ВРЕМЕННЫЕ ФАЙЛЫ
                    try
                    {

                        if (System.IO.File.Exists(textBox2.Text + "\\" + file.Name) == false)
                        {

                            try
                            {

                                if (checkBox4.Checked == true) // ЕСЛИ АВТОМАТИЧЕСКАЯ ПЕЧАТЬ ВКЛЮЧЕНА
                                {
                                    this.checkBox7.BeginInvoke((MethodInvoker)(() => checkBox7.Checked = true)); 
                                    ZipFile archive = ZipFile.Read(file.FullName);
                                    archive.ExtractAll(@"C:\Temp\Buber");
                                    LocalPrintServer printServer = new LocalPrintServer(PrintSystemDesiredAccess.AdministrateServer);
                                    PrintQueue pq = printServer.DefaultPrintQueue;
                                    PrintQueue printer = LocalPrintServer.GetDefaultPrintQueue();
                                    
                                    foreach (FileInfo fileZipSnp in dirTempInfo.GetFiles("*.SNP"))
                                    {
                                        flag = false;
                                        while (flag == false)
                                        {
                                            printer.Refresh();
                                            if (printer.IsBusy == true || printer.IsDoorOpened == true || printer.IsInError == true || printer.IsNotAvailable == true || printer.HasToner == false || printer.IsOffline == true || printer.IsPaperJammed == true || printer.IsOutOfPaper == true || printer.HasPaperProblem == true || printer.NeedUserIntervention == true)
                                            {
                                                MessageBox.Show("Проблемы с принтером", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                if (printer.NumberOfJobs >= 1)
                                                {
                                                    Thread.Sleep(1000);
                                                }
                                                else
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }

                                        ProcessStartInfo info = new ProcessStartInfo();
                                        info.Verb = "Printto";
                                        info.FileName = fileZipSnp.FullName;
                                        info.Arguments = "\"" + printDialog1.PrinterSettings.PrinterName + "\"";

                                        //MessageBox.Show(info.FileName.ToString());

                                        info.UseShellExecute = true;
                                        info.CreateNoWindow = true;
                                        info.WindowStyle = ProcessWindowStyle.Hidden;
                                        Process p = new Process();
                                        p.StartInfo = info;

                                        p.Start();
                                        p.WaitForExit();
                                        p.Close();
                                        p.Dispose();
                                    }

                                    foreach (FileInfo fileZipTxt in dirTempInfo.GetFiles("*.TXT"))
                                    {
                                        flag = false;
                                        while (flag == false)
                                        {
                                            printer.Refresh();
                                            if (printer.IsBusy == true || printer.IsDoorOpened == true || printer.IsInError == true || printer.IsNotAvailable == true || printer.HasToner == false || printer.IsOffline == true || printer.IsPaperJammed == true || printer.IsOutOfPaper == true || printer.HasPaperProblem == true || printer.NeedUserIntervention == true)
                                            {
                                                MessageBox.Show("Проблемы с принтером", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                if (printer.NumberOfJobs >= 1)
                                                {
                                                    Thread.Sleep(1000);
                                                }
                                                else
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }

                                        ProcessStartInfo info = new ProcessStartInfo();
                                        info.Verb = "Printto";
                                        info.FileName = fileZipTxt.FullName;
                                        info.Arguments = "\"" + printDialog1.PrinterSettings.PrinterName + "\"";

                                        //MessageBox.Show(info.FileName.ToString());

                                        info.UseShellExecute = true;
                                        info.CreateNoWindow = true;
                                        info.WindowStyle = ProcessWindowStyle.Hidden;
                                        Process p = new Process();
                                        p.StartInfo = info;

                                        p.Start();
                                        p.WaitForExit();
                                        p.Close();
                                        p.Dispose();
                                    }

                                    foreach (FileInfo fileZipDwgLsr in dirTempInfo.GetFiles("*lsr*.DWG"))
                                    {
                                        try
                                        {
                                            fileZipDwgLsr.Delete();
                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }

                                    foreach (FileInfo fileZipDwg in dirTempInfo.GetFiles("*.DWG"))
                                    {
                                        flag = false;
                                        while (flag == false)
                                        {
                                            printer.Refresh();
                                            if (printer.IsBusy == true || printer.IsDoorOpened == true || printer.IsInError == true || printer.IsNotAvailable == true || printer.HasToner == false || printer.IsOffline == true || printer.IsPaperJammed == true || printer.IsOutOfPaper == true || printer.HasPaperProblem == true || printer.NeedUserIntervention == true)
                                            {
                                                MessageBox.Show("Проблемы с принтером", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                if (printer.NumberOfJobs >= 1)
                                                {
                                                    Thread.Sleep(1000);
                                                }
                                                else
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }

                                        try
                                        {
                                            ProcessStartInfo startInfo = new ProcessStartInfo(textBox7.Text + "\\ACAD.EXE");
                                            startInfo.WindowStyle = ProcessWindowStyle.Minimized;

                                            startInfo.Arguments = "" + fileZipDwg.FullName + @" /b C:\Temp\main.scr /nologo";
                                            Process.Start(startInfo).WaitForExit();


                                            //Process.Start(textBox7.Text + "\\ACAD.EXE", " " + fileZipDwg.FullName + " /b C:\\temp\\test.scr /nologo").WaitForExit();

                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }

                                    this.checkBox7.BeginInvoke((MethodInvoker)(() => checkBox7.Checked = false)); 
                                }
                                
                                if (checkBox3.Checked == true) // ЕСЛИ АВТОМАТИЧЕСКАЯ РАСПАКОВКА ВКЛЮЧЕНА
                                {
                                    
                                    try
                                    {
                                        this.checkBox8.BeginInvoke((MethodInvoker)(() => checkBox8.Checked = true)); 
                                        ZipFile archive2 = ZipFile.Read(file.FullName);
                                        archive2.ExtractAll(textBox6.Text + "\\" + System.IO.Path.GetFileNameWithoutExtension(file.Name), ExtractExistingFileAction.OverwriteSilently);
                                        this.checkBox8.BeginInvoke((MethodInvoker)(() => checkBox8.Checked = false)); 
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }

                                


                                foreach (FileInfo fileDel in dirTempInfo.GetFiles())
                                {
                                    fileDel.Delete();
                                }
                            }
                            catch (Exception ex_print)
                            {
                                MessageBox.Show(ex_print.ToString());
                            }
                        }

                        // ИЩЕМ ПЕРЕЗАКИНУТЫЕ ЗАКАЗЫ

                        foreach (FileInfo file2 in dirInfo2.GetFiles(file.Name.ToString()))
                        {
                       
                            if (file.LastWriteTime.ToString() != file2.LastWriteTime.ToString() && file.Name == file2.Name && file.Exists == true)
                            {
                                try
                                {
                                    if (checkBox1.Checked == true)
                                    {
                                        File.Copy(file2.FullName, textBox2.Text + "\\" + Path.GetFileNameWithoutExtension(file2.FullName) + "_" + textBox4.Text + "_" + file2.LastWriteTime.ToShortDateString() + "-" + file2.LastWriteTime.ToString("HH-mm-ss") + file2.Extension.ToString(), true);
                                        File.Copy(file.FullName, textBox2.Text + "\\" + file.Name, true);
                                        notifyIcon1.ShowBalloonTip(60, "Бубер", "Был перезакинут файл " + file2.Name.ToString(), ToolTipIcon.Warning);

                                        DataGridViewCell firstCell = new DataGridViewTextBoxCell();
                                        DataGridViewCell secondCell = new DataGridViewTextBoxCell(); 
                                        DataGridViewCell thirdCell = new DataGridViewButtonCell();
                                        DataGridViewCell lastCell = new DataGridViewButtonCell();
                                        DataGridViewRow row = new DataGridViewRow();

                                        firstCell.Value = file2.Name.ToString();
                                        secondCell.Value = file.LastWriteTime.ToShortDateString() + " - " + file.LastWriteTime.ToShortTimeString();
                                        thirdCell.Value = "Печатать";
                                        lastCell.Value = "X";

                                        row.Cells.AddRange(firstCell, secondCell, thirdCell, lastCell);


                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Rows.Add(row)));

                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Refresh()));
                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Update()));
                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Parent.Refresh()));

                                        indx_tabl++;
                                        
                                    }
                                    else
                                    {
                                        File.Copy(file2.FullName, textBox2.Text + "\\" + Path.GetFileNameWithoutExtension(file2.FullName) + "_" + textBox4.Text + "_" + file2.Extension.ToString(), true);
                                        File.Copy(file.FullName, textBox2.Text + "\\" + file.Name, true);
                                        notifyIcon1.ShowBalloonTip(60, "Бубер", "Был перезакинут файл " + file2.Name.ToString(), ToolTipIcon.Warning);

                                        DataGridViewCell firstCell = new DataGridViewTextBoxCell();
                                        DataGridViewCell secondCell = new DataGridViewTextBoxCell();
                                        DataGridViewCell thirdCell = new DataGridViewButtonCell();
                                        DataGridViewCell lastCell = new DataGridViewButtonCell();
                                        DataGridViewRow row = new DataGridViewRow();

                                        firstCell.Value = file2.Name.ToString();
                                        secondCell.Value = file.LastWriteTime.ToShortDateString() + " - " + file.LastWriteTime.ToShortTimeString();
                                        thirdCell.Value = "Печатать";
                                        lastCell.Value = "X";

                                        row.Cells.AddRange(firstCell, secondCell, thirdCell, lastCell);


                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Rows.Add(row)));

                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Refresh()));
                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Update()));
                                        this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Parent.Refresh()));

                                        indx_tabl++;

                                        
                                    }
                                }
                                catch (Exception)
                                {
                                }
                            }
                            
                        }

                        this.checkBox5.BeginInvoke((MethodInvoker)(() => checkBox5.Checked = true));
                        try
                        {
                            File.Copy(file.FullName, textBox2.Text + "\\" + file.Name, false);
                        }
                        catch (Exception)
                        {
                        }
                        this.checkBox5.BeginInvoke((MethodInvoker)(() => checkBox5.Checked = false)); 

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }                   
                }
            }
            catch (Exception ex)
            {
                timer1.Enabled = false;
                MessageBox.Show(ex.Message);
            }
            //timer1.Enabled = true;
            threads_text--;
            threads_left++;
            this.linkLabel1.BeginInvoke((MethodInvoker)(() => this.linkLabel1.Text = threads_text.ToString()));
            this.linkLabel2.BeginInvoke((MethodInvoker)(() => this.linkLabel2.Text = threads_left.ToString()));
            this.checkBox6.BeginInvoke((MethodInvoker)(() => checkBox6.Checked = false)); 
            this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.Text = "Бубер - Отдыхаем"));
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.in_files;
            textBox2.Text = Properties.Settings.Default.out_files;
            textBox3.Text = Properties.Settings.Default.mask;
            textBox4.Text = Properties.Settings.Default.double_files;
            textBox5.Text = Properties.Settings.Default.cycle.ToString();
            textBox6.Text = Properties.Settings.Default.unpack_files;
            textBox7.Text = Properties.Settings.Default.autocad_dir;
            checkBox1.Checked = Properties.Settings.Default.check_double;
            checkBox2.Checked = Properties.Settings.Default.check_cycle;
            checkBox3.Checked = Properties.Settings.Default.unpack_auto;
            checkBox4.Checked = Properties.Settings.Default.autoprint;
            Properties.Settings.Default.program_dir = Environment.CurrentDirectory.ToString();
            timer1.Interval = Properties.Settings.Default.cycle * 1000;
            dataGridView1.Rows.Clear();

            if (checkBox2.Checked == true)
            {
                timer1.Enabled = true;
                if (Properties.Settings.Default.autostart == false) timer2.Enabled = true;
            }
            timer3.Enabled = true;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            Properties.Settings.Default.in_files = textBox1.Text;
            Properties.Settings.Default.out_files = textBox2.Text;
            Properties.Settings.Default.unpack_files = textBox6.Text;
            Properties.Settings.Default.autoprint = checkBox4.Checked;
            Properties.Settings.Default.unpack_auto = checkBox3.Checked;
            Properties.Settings.Default.mask = textBox3.Text;
            Properties.Settings.Default.double_files = textBox4.Text;
            Properties.Settings.Default.cycle = int.Parse(textBox5.Text);
            Properties.Settings.Default.check_double = checkBox1.Checked;
            Properties.Settings.Default.check_cycle = checkBox2.Checked;
            Properties.Settings.Default.autostart = false;
            Properties.Settings.Default.autocad_dir = textBox7.Text;
            Properties.Settings.Default.Save();
            Environment.Exit(0);
        }

        private void textBox5_MouseDown(object sender, MouseEventArgs e)
        {
            textBox5.SelectAll();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath.ToString() != "folderBrowserDialog1") textBox1.Text = folderBrowserDialog1.SelectedPath.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath.ToString() != "folderBrowserDialog1") textBox2.Text = folderBrowserDialog1.SelectedPath.ToString();
        }

        private void textBox3_MouseDown(object sender, MouseEventArgs e)
        {
            textBox3.SelectAll();
        }

        private void textBox4_MouseDown(object sender, MouseEventArgs e)
        {
            textBox4.SelectAll();
        }

        private void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            textBox1.SelectAll();
        }

        private void textBox2_MouseDown(object sender, MouseEventArgs e)
        {
            textBox2.SelectAll();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
                if (checkBox2.Checked == true)
                {
                    this.notifyIcon1.ShowBalloonTip(15, "Бубер", "Я теперь буду работать тут", ToolTipIcon.Info);
                    timer1.Enabled = true;
                    timer2.Enabled = true;
                }
                else
                {
                    if (threads_text < 1)
                    {
                        period_timer.Enabled = true;
                        all_period_timer.Enabled = true;
                        Thread myThread = new Thread(start_copy);
                        myThread.IsBackground = false;                        
                        myThread.Start();
                    }
                }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
                if (checkBox2.Checked == true)
                {
                    if (threads_text < 1)
                    {
                        period_timer.Enabled = true;
                        period_timer.Start();
                        all_period_timer.Enabled = true;
                        linkLabel5.Text = "00:00:00";
                        Thread myThread = new Thread(start_copy);
                        myThread.IsBackground = false;
                        myThread.Start();
                    }
                }
                else
                {
                    timer1.Enabled = false;
                }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            timer1.Enabled = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Супер-мега-копирэн программа БУБЕР!!!" + "\n" + "\n" + "Написал: Костюков И.К." + "\n" + "Тестировали: Бабурин М.А. и Иванов С.В.", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            Properties.Settings.Default.in_files = textBox1.Text;
            Properties.Settings.Default.out_files = textBox2.Text;
            Properties.Settings.Default.unpack_files = textBox6.Text;
            Properties.Settings.Default.autoprint = checkBox4.Checked;
            Properties.Settings.Default.unpack_auto = checkBox3.Checked;
            Properties.Settings.Default.mask = textBox3.Text;
            Properties.Settings.Default.double_files = textBox4.Text;
            Properties.Settings.Default.cycle = int.Parse(textBox5.Text);
            Properties.Settings.Default.check_double = checkBox1.Checked;
            Properties.Settings.Default.check_cycle = checkBox2.Checked;
            Properties.Settings.Default.autostart = false;
            Properties.Settings.Default.Save();
            Environment.Exit(0);
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.autostart = true;
            this.Visible = true;
            
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Properties.Settings.Default.autostart = true;
            this.Visible = true;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.check_cycle = checkBox2.Checked;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.check_double = checkBox1.Checked;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Properties.Settings.Default.double_files = textBox4.Text;
            }
            catch (Exception)
            {
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Properties.Settings.Default.mask = textBox3.Text;
            }
            catch (Exception)
            {
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            
                Properties.Settings.Default.autostart = true;
                this.Visible = false;
                timer2.Enabled = false;
            
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                //MessageBox.Show(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
                DirectoryInfo dirInfo2 = new DirectoryInfo(textBox2.Text);
                DirectoryInfo dirTempInfo = new DirectoryInfo(@"C:\Temp\Buber");
                foreach (FileInfo file in dirInfo2.GetFiles(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString()))
                {
                    // РАСПАКОВЫВАЕМ И ПЕЧАТАЕМ И УДАЛЯЕМ ВРЕМЕННЫЕ ФАЙЛЫ
                        try
                        {
                            if (System.IO.File.Exists(textBox2.Text + "\\" + file.Name) == true)
                            {
                                try
                                {
                                    bool flag = false;

                                    LocalPrintServer printServer = new LocalPrintServer(PrintSystemDesiredAccess.AdministrateServer);
                                    PrintQueue pq = printServer.DefaultPrintQueue;

                                    PrintQueue printer = LocalPrintServer.GetDefaultPrintQueue();

                                    ZipFile archive = ZipFile.Read(file.FullName);
                                    archive.ExtractAll(@"C:\Temp\Buber");

                                    foreach (FileInfo fileZipSnp in dirTempInfo.GetFiles("*.SNP"))
                                    {
                                        flag = false;
                                        while (flag == false)
                                        {
                                            printer.Refresh();
                                            if (printer.IsBusy == true || printer.IsDoorOpened == true || printer.IsInError == true || printer.IsNotAvailable == true || printer.HasToner == false || printer.IsOffline == true || printer.IsPaperJammed == true || printer.IsOutOfPaper == true || printer.HasPaperProblem == true || printer.NeedUserIntervention == true)
                                            {
                                                MessageBox.Show("Проблемы с принтером", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                if (printer.NumberOfJobs >= 1)
                                                {
                                                    Thread.Sleep(1000);
                                                }
                                                else
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }

                                        ProcessStartInfo info = new ProcessStartInfo();
                                        info.Verb = "Printto";
                                        info.FileName = fileZipSnp.FullName;
                                        info.Arguments = "\"" + printDialog1.PrinterSettings.PrinterName + "\"";
                                        info.UseShellExecute = true;
                                        info.CreateNoWindow = true;
                                        info.WindowStyle = ProcessWindowStyle.Hidden;
                                        Process p = new Process();
                                        p.StartInfo = info;
                                        p.Start();
                                        p.WaitForExit();
                                        p.Close();
                                        p.Dispose();
                                    }

                                    foreach (FileInfo fileZipTxt in dirTempInfo.GetFiles("*.TXT"))
                                    {
                                        flag = false;
                                        while (flag == false)
                                        {
                                            printer.Refresh();
                                            if (printer.IsBusy == true || printer.IsDoorOpened == true || printer.IsInError == true || printer.IsNotAvailable == true || printer.HasToner == false || printer.IsOffline == true || printer.IsPaperJammed == true || printer.IsOutOfPaper == true || printer.HasPaperProblem == true || printer.NeedUserIntervention == true)
                                            {
                                                MessageBox.Show("Проблемы с принтером", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                if (printer.NumberOfJobs >= 1)
                                                {
                                                    Thread.Sleep(1000);
                                                }
                                                else
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }

                                        ProcessStartInfo info = new ProcessStartInfo();
                                        info.Verb = "Printto";
                                        info.FileName = fileZipTxt.FullName;
                                        info.Arguments = "\"" + printDialog1.PrinterSettings.PrinterName + "\"";
                                        info.UseShellExecute = true;
                                        info.CreateNoWindow = true;
                                        info.WindowStyle = ProcessWindowStyle.Hidden;
                                        Process p = new Process();
                                        p.StartInfo = info;

                                        p.Start();
                                        p.WaitForExit();
                                        p.Close();
                                        p.Dispose();
                                    }

                                    if (checkBox3.Checked == true) // ЕСЛИ АВТОМАТИЧЕСКАЯ РАСПАКОВКА ВКЛЮЧЕНА
                                    {
                                        try
                                        {
                                            ZipFile archive2 = ZipFile.Read(file.FullName);
                                            archive2.ExtractAll(textBox6.Text + "\\" + System.IO.Path.GetFileNameWithoutExtension(file.Name), ExtractExistingFileAction.OverwriteSilently);
                                        }
                                        catch (Exception)
                                        {

                                        }
                                    }

                                    
                                    foreach (FileInfo fileZipDwgLsr in dirTempInfo.GetFiles("*lsr*.DWG"))
                                    {
                                        try
                                        {
                                            fileZipDwgLsr.Delete();
                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }

                                    foreach (FileInfo fileZipDwg in dirTempInfo.GetFiles("*.DWG"))
                                    {
                                        flag = false;
                                        while (flag == false)
                                        {
                                            printer.Refresh();
                                            if (printer.IsBusy == true || printer.IsDoorOpened == true || printer.IsInError == true || printer.IsNotAvailable == true || printer.HasToner == false || printer.IsOffline == true || printer.IsPaperJammed == true || printer.IsOutOfPaper == true || printer.HasPaperProblem == true || printer.NeedUserIntervention == true)
                                            {
                                                MessageBox.Show("Проблемы с принтером", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                Thread.Sleep(1000);
                                            }
                                            else
                                            {
                                                if (printer.NumberOfJobs >= 1)
                                                {
                                                    Thread.Sleep(1000);
                                                }
                                                else
                                                {
                                                    flag = true;
                                                }
                                            }
                                        }

                                        try
                                        {
                                            ProcessStartInfo startInfo = new ProcessStartInfo(textBox7.Text + "\\ACAD.EXE");
                                            startInfo.WindowStyle = ProcessWindowStyle.Minimized;

                                            startInfo.Arguments = "" + fileZipDwg.FullName + @" /b C:\Temp\main.scr /nologo";
                                            Process.Start(startInfo).WaitForExit();


                                            //Process.Start(textBox7.Text + "\\ACAD.EXE", " " + fileZipDwg.FullName + " /b C:\\temp\\test.scr /nologo").WaitForExit();

                                        }
                                        catch (Exception)
                                        {
                                        }
                                    }
                                    foreach (FileInfo fileDel in dirTempInfo.GetFiles())
                                    {
                                        fileDel.Delete();
                                    }
                                }
                                catch (Exception)
                                {
                                }
                            }

                        }
                        catch (Exception)
                        {
                        }
                }
                this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Rows.RemoveAt(e.RowIndex)));
                dataGridView1.Refresh();
                dataGridView1.Update();
                dataGridView1.Parent.Refresh();
                indx_tabl--;
            }
            if (e.ColumnIndex == 3)
            {
                this.dataGridView1.BeginInvoke((MethodInvoker)(() => this.dataGridView1.Rows.RemoveAt(e.RowIndex)));
                dataGridView1.Refresh();
                dataGridView1.Update();
                dataGridView1.Parent.Refresh();
                indx_tabl--;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath.ToString() != "folderBrowserDialog1") textBox6.Text = folderBrowserDialog1.SelectedPath.ToString();
        }

        private void textBox6_MouseDown(object sender, MouseEventArgs e)
        {
            textBox6.SelectAll();
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            Properties.Settings.Default.in_files = textBox1.Text;
            Properties.Settings.Default.out_files = textBox2.Text;
            Properties.Settings.Default.unpack_files = textBox6.Text;
            Properties.Settings.Default.autoprint = checkBox4.Checked;
            Properties.Settings.Default.unpack_auto = checkBox3.Checked;
            Properties.Settings.Default.mask = textBox3.Text;
            Properties.Settings.Default.double_files = textBox4.Text;
            Properties.Settings.Default.cycle = int.Parse(textBox5.Text);
            Properties.Settings.Default.check_double = checkBox1.Checked;
            Properties.Settings.Default.check_cycle = checkBox2.Checked;
            Properties.Settings.Default.autostart = false;
            Properties.Settings.Default.autocad_dir = textBox7.Text;
            Properties.Settings.Default.Save();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            int exitCode;
            using (var process = new Process())
            {
                var startInfo = process.StartInfo;
                startInfo.FileName = "cmd";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Verb = "runas";
                startInfo.Arguments = string.Format("/c net stop spooler");

                process.Start();
                process.WaitForExit();

                exitCode = process.ExitCode;

                process.Close();
                Thread.Sleep(1000);
                startInfo.FileName = "cmd";
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.Verb = "runas";
                startInfo.Arguments = string.Format("/c net start spooler");

                process.Start();
                process.WaitForExit();

                exitCode = process.ExitCode;

                process.Close();

                MessageBox.Show("Служба перезагружена", "Бубер", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void period_timer_Tick(object sender, EventArgs e)
        {
            ++p_sec;
            if (p_sec != 0) this.linkLabel5.BeginInvoke((MethodInvoker)(() => this.linkLabel5.Text = String.Format("{0:00}:{1:00}:{2:00}", p_sec / 3600, p_sec / 60 % 60, p_sec % 60).ToString()));
            if (p_sec == 0) this.linkLabel5.BeginInvoke((MethodInvoker)(() => this.linkLabel5.Text = "00:00:00"));
        }

        private void linkLabel1_TextChanged(object sender, EventArgs e)
        {
            if (threads_text == 0)
            {
                this.period_timer.Enabled = false;
                this.all_period_timer.Enabled = false;
                p_sec = 0;
            }
        }

        private void all_period_timer_Tick(object sender, EventArgs e)
        {
            ++a_sec;
            if (a_sec != 0) this.linkLabel8.BeginInvoke((MethodInvoker)(() => this.linkLabel8.Text = String.Format("{0:00}:{1:00}:{2:00}", a_sec / 3600, a_sec / 60 % 60, a_sec % 60).ToString()));
            if (a_sec == 0) this.linkLabel8.BeginInvoke((MethodInvoker)(() => this.linkLabel8.Text = "00:00:00"));
        }

        private void button10_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath.ToString() != "folderBrowserDialog1") textBox7.Text = folderBrowserDialog1.SelectedPath.ToString();
            Properties.Settings.Default.autocad_dir = textBox7.Text;
        }

        private void textBox7_MouseDown(object sender, MouseEventArgs e)
        {
            textBox7.SelectAll();
        }
    }
}
