using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Threading;
using SpeechLib;
using System.Diagnostics;
using System.Data.SqlClient;
using mshtml;
using WatiN.Core;
using WatiN.Core.Native.Windows;
using System.Runtime.InteropServices;


namespace speechRecognition
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        SqlConnection connect = new SqlConnection("Data Source=.\\FURKAN;Initial Catalog=speechRecognition;Integrated Security=True");
        
        private SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine();
        SpVoice spVoice = new SpVoice();
        SpeechVoiceSpeakFlags svsp = new SpeechVoiceSpeakFlags();
        
        string[] wordsAll;
        int count = 0;
        int wordCount = 0;
        bool ExLearnLink = false;

        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const int MOUSEEVENTF_LEFTUP = 0x0004;

        void words()
        {
            connect.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connect;

            cmd.CommandText = "select count(*) from words";
            wordCount = Convert.ToInt32(cmd.ExecuteScalar());
            wordsAll = new string[wordCount];
            cmd.ExecuteNonQuery();

            cmd.CommandText = "select * from words";
            
            SqlDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                wordsAll[count]  = dr["word"].ToString();
                count++;
            }
            connect.Close();
        }

        public void LoadGrammer()
        {
            Choices choices = new Choices();
            for (int i = 0; i < wordsAll.Length; i++)
            {
                choices.Add(new string[] { wordsAll[i] });
            }
            GrammarBuilder grammarBuilder = new GrammarBuilder(choices);

            grammarBuilder.Culture = System.Globalization.CultureInfo.GetCultureInfoByIetfLanguageTag("en-US");
            Grammar grammar = new Grammar(grammarBuilder);
            recognizer.LoadGrammar(grammar);
            spVoice.Speak("Welcome Furkan", svsp);       

        }

        private void recognizer_SpeechDetected(object sender, SpeechDetectedEventArgs e)
        {

        }

        private void recognizer_SpeechRecognitionRejected(object sender, SpeechRecognitionRejectedEventArgs e)
        {
            spVoice.Speak("i can't understand you please repeat it?", svsp);
        }

        public void DoMouseClick()
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, System.Windows.Forms.Control.MousePosition.X, System.Windows.Forms.Control.MousePosition.Y, 0, 0);
        }


        IE ie;
        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text == "open league of legends")
            {
                Process.Start("D:\\Riot Games\\League of Legends\\LeagueClient.exe");
                spVoice.Speak("I opened League of Legends");
            }
            else if(e.Result.Text == "Say my name"){
                spVoice.Speak("Furkan YUKSEL");
            }
            else if (e.Result.Text == "what is your name")
            {
                spVoice.Speak("My name is Ellie");
            }
            else if (e.Result.Text == "minimize")
            {
                this.WindowState = FormWindowState.Minimized;
            }
            else if(e.Result.Text == "close")
            {
                spVoice.Speak("Im out");
                recognizer.Dispose();
                Application.Exit();
            }
        }

        private void recognizer_SpeechRecognized2(object sender, SpeechRecognizedEventArgs ev)
        {
            MessageBox.Show(ev.Result.Text.ToString());
            for (int i = 0; i < wordsAll.Length; i++)
            {
			    if (ev.Result.Text == "search" + " " + wordsAll[i])
                {
                    spVoice.Speak(ev.Result.Text.ToString());
                    ie = new IE("http://www.google.com/");
                    if (ie != null)
                    {
                        ie.ShowWindow(NativeMethods.WindowShowStyle.Maximize);
                        ie.TextField(Find.ById("lst-ib")).Focus();
                    }
                    if (ie != null && ie.NativeBrowser != null)
                    {
                        ie.TextField(Find.ById("lst-ib")).Value += wordsAll[i] + " ";
                    }
                    Cursor.Position = new Point(Cursor.Position.X + 200, Cursor.Position.Y + 200);
                    DoMouseClick();
                    ie.Button(Find.ById("mKlEF")).Click();
                }
			}
            if (ev.Result.Text == "exit google")
            {
                if (ie != null)
                {
                    ExLearnLink = false;
                    ie.Close();
                }
            }
        }

        private void recognizer_RecognizeCompleted(object sender, RecognizeCompletedEventArgs e)
        {
            recognizer.RecognizeAsync();
        }

        private void StartRecognition()
        {
            recognizer.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(recognizer_SpeechDetected);

            recognizer.SpeechRecognitionRejected += new EventHandler<SpeechRecognitionRejectedEventArgs>(recognizer_SpeechRecognitionRejected);

            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);

            recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized2);

            recognizer.RecognizeCompleted += new EventHandler<RecognizeCompletedEventArgs>(recognizer_RecognizeCompleted);

            Thread t1 = new Thread(delegate()
            {
                recognizer.SetInputToDefaultAudioDevice();
                recognizer.RecognizeAsync(RecognizeMode.Single);
            });
            t1.Start();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            words();
            LoadGrammer();
            StartRecognition();
        }
    }
}
