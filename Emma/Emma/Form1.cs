using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using System.Xml;

namespace Emma
{
    public partial class Form1 : Form
    {

        SpeechSynthesizer s = new SpeechSynthesizer();
        Boolean wake = false;

        String name = "Lasith Tharindu";
        Boolean var1 = true;
        String temp;
        String condition;
        Choices list = new Choices();

        public Boolean search = false;

        public Form1()
        {

            SpeechRecognitionEngine rec = new SpeechRecognitionEngine();

            list.Add(new String[] { "Hello",
                "How are you",
                "What time is it",
                "What day is it",
                "Open Google",
                "wake up","sleep",
                "restart","update",
                "open word",
                "close word",
                "whats the weather like",
                "what is the tempereature",
                "hey Emma",
                "tell me a joke",
                "minimize","maximize",
                 "unminimize",
                "play some music",
                "pause",
                "open Windows Media Player",
                "close Windows Media Player",
                "next",
                "Previous",
                "whats my name",
                "do i like fastfood",
                "search",
                "what can you do",
                 "What is the date",
                "What are the days for weeks",
                "Who are you",
                 "close this",
                "stop"});

            Grammar gr = new Grammar(new GrammarBuilder(list));
            try
            {
                rec.RequestRecognizerUpdate();
                rec.LoadGrammar(gr);
                rec.SpeechRecognized += rec_SpeachRecognized;
                rec.SetInputToDefaultAudioDevice();
                rec.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch { return; }


            s.SelectVoiceByHints(VoiceGender.Female);
           // s.Speak("Hello ,My name is Emma");
            InitializeComponent();
        }

        public String GetWeather(String input)
        {
            String query = String.Format("https://query.yahooapis.com/v1/public/yql?q=select * from weather.forecast where woeid in (select woeid from geo.places(1) where text='Gampaha, or')&format=xml&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys");
            XmlDocument wData = new XmlDocument();
            try
            {
                wData.Load(query);

            }
            catch
            {
                MessageBox.Show("No Internet Connection! Please try again late");
                return "no Internet";
            }
            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("query").SelectSingleNode("results").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("query/results/channel");
            try
            {
                temp = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;
                condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;
                if (input == "temp")
                {
                    return temp;
                }
               
                if (input == "cond")
                {
                    return condition;
                }
            }
            catch
            {
                return "Error Reciving data";
            }
            return "error";
        }
        public static void killprog(String s)
        {
            Process[] procs = null;
            try
            {
                procs = Process.GetProcessesByName(s);
                Process prog = procs[0];

                if (!prog.HasExited)
                {
                    prog.Kill();
                }
            }
            finally
            {
                if(procs != null)
                {
                    foreach(Process p in procs)
                    {
                        p.Dispose();
                    }
                }
            }
            procs = null;
        }

        public void say(String h)
        {
            s.Speak(h);
           // s.SpeakAsync(h);
            textBox2.AppendText(h + "\n");
        }

        String[] greeting = new String[3] { "Hi", "Hello", "Hi,How are you?"};

        public String greeting_action()
        {
            Random r = new Random();
            return greeting[r.Next(3)];
        }
        public void restart()
        {
              Process.Start(@"E:\Emma\Emma.exe");
             Environment.Exit(0);
        }
        private void rec_SpeachRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            String r = e.Result.Text;

            //if (search)
            //{
            //    Process.Start("https://google.com/search?q=" + r);
            //}
            if (r== "hey Emma")
            {
                wake = true;
                say(greeting_action());
            }
            if (r == "wake up")
            {
                wake = true;
                Status.Text = "Status: Wake";
            }
            if (r == "sleep")
            {
                wake = false;
                Status.Text = "Status: Sleep Mode";
            }

            if (wake == true)
            {
                if(r.StartsWith("search"))
                {
                    String search = r.Replace("google", "");
                    System.Diagnostics.Process.Start("http://google.com/search?q=" + search);
                }

                if (r == "stop")
                {
                    s.SpeakAsyncCancelAll();
                }

                if (r == "do i like fastfood")
                {
                    if (!var1)
                        say("No ," +name+ "You dont like fastfood.");
                    if (var1 )
                        say("yes," + name +"You love fastfood");
                }
                if (r=="whats my name")
                {
                    say("Your name is " + name);
                }
                if (r == "Previous")
                {
                    SendKeys.Send("^{LEFT}");
                }
                if (r == "next")
                {
                    SendKeys.Send("^{RIGHT}");
                }
                if(r== "open Windows Media Player")
                {
                    say("Opeing Windows Media Player");
                    Process.Start(@"C:\Program Files (x86)\Windows Media Player\wmplayer.exe");
                }
                if (r == "play some music" || r == "pause")
                {
                    SendKeys.Send("");
                }
                if (r == "what can you do")
                {
                    say("I can read email, weather report, i can search web for you, i can fix and tell you about your appointments, anything that you need like a persoanl assistant do, you can ask me question i will reply to you");
                }
                if (r== "Tell me a joke")
                {
                    SendKeys.Send("");
                }
                if (r == "unminimize")
                {
                    this.WindowState = FormWindowState.Normal;
                }
                if (r == "minimize")
                {
                    this.WindowState = FormWindowState.Minimized;
                }
                if (r == "maximize")
                {
                    this.WindowState = FormWindowState.Maximized;
                }
                if (r =="close this")
                {
                    Application.Exit();
                }
                if (r=="whats the weather like")
                {

                    say("The Sky is," + GetWeather("cond") + ",");
                }
                if (r == "what is the tempereature")
                {
                    say("It is," + GetWeather("temp") + "degrees.");
                }
                if(r== "Who are you")
                {
                    say("I am your personal assistant, you can ask question i will reply to you");
                }
                if (r=="open word")
                {
                    Process.Start(@"C:\Program Files\Microsoft Office\Office16\WINWORD.exe");
                }
                if(r=="close word")
                {
                    killprog("WINWORD.exe");
                }
                if(r=="restart" || r == "update")
                {
                    restart();
                }
                if (r== "close Windows Media Player")
                {
                    killprog("wmplayer.bin");
                }

                if (r == "Hello")
                {
                    say("hi");
                }

                if (r == "What time is it")
                {
                    say(DateTime.Now.ToString("h:mm tt"));
                }
                if (r == "What is the date")
                {
                    say(DateTime.Now.ToString("M/d/yyyy"));
                }
                if(r=="what day is it")
                {
                    say(DateTime.Now.ToString());
                }
                if(r == "What are the days for weeks")
                {
                    say(Day.Monday.ToString());
                    say(Day.Tuesday.ToString());
                    say(Day.Wednesday.ToString());
                    say(Day.Thursday.ToString());
                    say(Day.Friday.ToString());
                    say(Day.Saturday.ToString());
                    say(Day.Sunday.ToString());
                }
                if (r == "How are you")
                {
                    say("Greate, and You?");
                }
                if (r == "Open Google")
                {
                    Process.Start("http://www.google.com");
                }
            }
            textBox1.AppendText(r + "\n");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
