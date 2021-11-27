using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.Structure;
using System.Threading;

        
namespace Face
{
    public partial class Form1 : Form
    {
        private object lockObject = new object();
        static readonly CascadeClassifier _cascadeClassifier = new CascadeClassifier("haarcascade_frontalface_alt_tree.xml");


        public Form1()
        {
            InitializeComponent();
       

        }
        

        private void button1_Click(object sender, EventArgs e)
        {
             /*Thread th1 = new Thread(new ThreadStart(khuonmat));
             Control.CheckForIllegalCrossThreadCalls = false;
             th1.IsBackground = true;
             th1.Start();
             Thread.Sleep(2000);
             //khuonmat();
           Thread t = new Thread((ThreadStart)(() => {
                khuonmat();
                
            }));
            // Run your code from a thread that joins the STA Thread
            t.SetApartmentState(ApartmentState.STA);
            //Control.CheckForIllegalCrossThreadCalls = false;
            Thread.Sleep(100);
            t.Start();
            //t.Join();

            //  khuonmat();
            */
        }

        void khuonmat()
        {
            lock (lockObject)
            {

                /*using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "JPEG|*.jpg" })
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                       // pictureBox1.Image = Image.FromFile(openFileDialog.FileName);



                        Bitmap img = new Bitmap(Image.FromFile(openFileDialog.FileName));
                        */

                using (var fbd = new FolderBrowserDialog())
                {
                    DialogResult result = fbd.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    {
                        string[] files = Directory.GetFiles(fbd.SelectedPath, "*.jpg");
                        

                        {
                            int i = 0;

                            while (i < files.Length)
                            {
                                Bitmap img = new Bitmap(Image.FromFile(files[i]));

                                Image<Bgr, byte> grayframe = new Image<Bgr, byte>(img);
                                // pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                                //pictureBox1.Image = img

                                // string[] FileList = Directory.GetFiles("C:/Users/Admin/Pictures/Saved Pictures", "*", SearchOption.AllDirectories);
                                // label1.Text = openFileDialog.FileName;
                                //Hiển thị danh sách file
                                // foreach (var item in FileList)
                                //{
                                // ListViewItem item1 = new ListViewItem();
                                //  listView1.Items.Add(item);
                                //     listBox1.Items.Add(item);
                                //  }
                                listBox1.Invoke(new MethodInvoker(() =>
                               {
                                   if (listBox1.Items != null)
                                   {
                                       listBox1.Items.Clear();
                                       foreach (var item in files)
                                       {
                                           listBox1.Items.Add(item);
                                       }
                                   }
                               }));



                                Rectangle[] faces = _cascadeClassifier.DetectMultiScale(grayframe, 1.1, 1);
                                if (faces.Length > 0)
                                {
                                    Bitmap BmpInput = grayframe.ToBitmap();
                                    Bitmap ExtractedFace;
                                    Graphics FaceCanvas;
                                 
                                    foreach (var face in faces)
                                    {
                                        grayframe.Draw(face, new Bgr(Color.Blue), 4);
                                        ExtractedFace = new Bitmap(face.Width, face.Height);
                                        FaceCanvas = Graphics.FromImage(ExtractedFace);
                                        FaceCanvas.DrawImage(BmpInput, 0, 0, face, GraphicsUnit.Pixel);
                                        if (face.Width < 100) { return; }
                                        int w = face.Width;
                                        int h = face.Height;
                                        int x = face.X;
                                        int y = face.Y;
                                        int r = Math.Max(250, 250) / 2;
                                        int centerx = x + w / 2;
                                        int centery = y + h / 2;
                                        int nx = (int)(centerx - r);
                                        int ny = (int)(centery - r);
                                        int nr = (int)(r * 5);
                                        double zoomFactor = (double)w / (double)face.Width;
                                        Size newSize = new Size((int)(img.Width * zoomFactor), (int)(img.Height * zoomFactor));
                                        Bitmap bmp = new Bitmap(img, newSize);
                                        Invoke((MethodInvoker)(delegate ()
                                        {
                                            if (this.panel1.Controls["pictureBox0"] == null)
                                            {
                                                PictureBox pic = new PictureBox();
                                                pic.BackColor = System.Drawing.SystemColors.ActiveCaption;
                                                pic.Location = new System.Drawing.Point(253, 4);
                                                pic.Name = "pictureBox0";
                                                pic.Size = new System.Drawing.Size(93, 106);
                                                pic.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
                                                pic.TabIndex = 2;
                                                pic.TabStop = false;
                                                // pic.Text = files[i];
                                                this.panel1.Controls.Add(pic);
                                            }
                                            if (this.Controls["pictureBox" + (i+1)] == null)

                                            TaopicBenDuoi((PictureBox)this.panel1.Controls["pictureBox" + i], "pictureBox" + (i + 1));
                                        else if (i >= files.Length)
                                           this.Controls["pictureBox" + (i + 1)].Dispose();
                                        }));

                                        void TaopicBenDuoi(PictureBox TextBoxBenTren, string picName)
                                        {

                                            PictureBox tbx = new PictureBox();
                                            var imgextract = CropImage(img, nx + 4, ny - 25, 248, 340);
                                         
                                            tbx.Top = TextBoxBenTren.Bottom + 1;
                                            tbx.Left = TextBoxBenTren.Left;
                                            tbx.Width = TextBoxBenTren.Width;
                                            tbx.Height = TextBoxBenTren.Height;
                                            tbx.Name = picName;
                                            TextBoxBenTren.Image= imgextract;
                                           TextBoxBenTren.SizeMode = PictureBoxSizeMode.StretchImage;
                                            this.panel1.Controls.Add(tbx);
                                            int n = files.Length;
                                            int pt = 100 / n;
                                            Thread.Sleep(2000);
                                               // progressBar1.Step =n;
                                                progressBar1.Increment(0);
                                            if (TextBoxBenTren.Image != null)
                                            {

                                                progressBar1.Value  = progressBar1.Value + pt;
                                            }
                                                label2.Text = progressBar1.Value.ToString() + "%";
                                            if (progressBar1.Value >= 90 && i==files.Length-1)
                                            {
                                                progressBar1.Value = progressBar1.Value=100;                                      
                                            }
                                            if (progressBar1.Value == 100)
                                            {
                                                label2.Text = "Done!";
                                                //progressBar1.Value = 1;
                                            }


                                        }
                                        
                                        Invoke((MethodInvoker)(delegate ()
                                        {
                                         
                                           
                                            
                                            if (this.panel1.Controls["textBox0"] == null)
                                            {
                                                TextBox tx = new TextBox();
                                                tx.Location = new System.Drawing.Point(15, 17);
                                                tx.Margin = new System.Windows.Forms.Padding(200);
                                                tx.Name = "textBox0";
                                                tx.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
                                                tx.Size = new System.Drawing.Size(222, 22);
                                                tx.TabIndex = 18;
                                                tx.Text = files[i];
                                                tx.Top = 50;
                                                this.panel1.Controls.Add(tx);
                                                
                                            }
                                            else if (this.Controls["textBox" +i] == null)
                                            {
                                                TaoTextBoxBenDuoi((TextBox)this.panel1.Controls["textBox" + (i-1)], "textBox" + i);
                                                this.panel1.Controls["textBox" + i].Text = files[i];
                                            }
                                            var index = dataGridView1.Rows.Add();
                                            dataGridView1.Rows[i].Cells["Column1"].Value = this.panel1.Controls["textBox" + i].Text;
                                            dataGridView1.Rows[i].Cells["Column2"].Value = files[i];

                                        }));
                                       
                                        void TaoTextBoxBenDuoi(TextBox TextBoxBenTren, String TextBoxName)
                                        {

                                            TextBox tbx = new TextBox();
                                            tbx.Top = TextBoxBenTren.Bottom+90;
                                            tbx.Left = TextBoxBenTren.Left;
                                            tbx.Width = TextBoxBenTren.Width;
                                            tbx.Name = TextBoxName;
                                            tbx.ScrollBars = TextBoxBenTren.ScrollBars;
                                           

                                            this.panel1.Controls.Add(tbx);
                                        }

                                          

                                        
                                        /* if (pictureBox1.Image == null)
                                         {
                                             pictureBox1.Image = (Image)bmp;
                                             var imgextract = CropImage(pictureBox1.Image, nx + 4, ny - 25, 248, 340);
                                             pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                                             pictureBox1.Image = imgextract;
                                         }*/


                                    }

                                    

                                }

                                i++;
                            }
                        }
                    }
                }
            } 

        }
        
        public Bitmap CropImage(Image source, int x, int y, int width, int height)
        {
            Rectangle crop = new Rectangle(x, y, width, height);

            var bmp = new Bitmap(crop.Width, crop.Height);
            using (var gr = Graphics.FromImage(bmp))
            {
                Invoke((MethodInvoker)(delegate ()
                {
                    gr.DrawImage(source, new Rectangle(0, 0, bmp.Width, bmp.Height), crop, GraphicsUnit.Pixel);
                }));

            }
            return bmp;
        }


        private void button2_Click(object sender, EventArgs e)
        {
        
           

            SaveFileDialog save = new SaveFileDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ToExcel(dataGridView1, saveFileDialog1.FileName);
            }

        }
      
        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!timer1.Enabled)
            {
                timer1.Start();
            }
            else
                timer1.Stop();

            Thread t = new Thread((ThreadStart)(() => {
                khuonmat();

            }));
            // Run your code from a thread that joins the STA Thread
            t.SetApartmentState(ApartmentState.STA);
            //Control.CheckForIllegalCrossThreadCalls = false;
            Thread.Sleep(100);
            t.Start();
           
           // demo();
           
        }
        private void demo(int n)
        {
            while (n <= 100)
            {
              

                //timer1.Interval =phantram
                switch (n)
                {
                    case 5:
                        timer1.Start();
                        // progressBar1.PerformStep();
                        progressBar1.Value = progressBar1.Value+5;
                        //timer1.Stop();
                        break;
                   
                    case 10:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 15:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;

                    case 20:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 25:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;

                    case 30:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 35:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;

                    case 40:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 45:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 50:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 55:
                        timer1.Start();
                        // progressBar1.PerformStep();
                        progressBar1.Value = progressBar1.Value + 5;
                        //timer1.Stop();
                        break;

                    case 60:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 65:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;

                    case 70:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 75:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;

                    case 80:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 85:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;

                    case 90:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 95:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;
                    case 100:
                        timer1.Start();
                        progressBar1.Value = progressBar1.Value + 5;
                        break;

                }

                     n++;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            
           

            // progress bar tăng lên từ từ
            //progressBar1.Value++;
            /*  Thread.Sleep(2000);
             if(progressBar1.Value<100)
              {
                  progressBar1.Value++;
                  label2.Text = progressBar1.Value.ToString() + "%";
              }
              else
              {
                  label2.Text = "Done!";
              }*/
        }
        private void ToExcel(DataGridView dataGridView1, string fileName)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                workbook = excel.Workbooks.Add(Type.Missing);

                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["Sheet1"];
                worksheet.Name = "HOC LAP TRINH C#";

                // export header
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }

                // export content
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                // save workbook
                workbook.SaveAs(fileName);
                workbook.Close();
                excel.Quit();
                MessageBox.Show("Export successful.!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                workbook = null;
                worksheet = null;
            }
        }
       
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
