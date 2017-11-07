using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using excel = Microsoft.Office.Interop.Excel; // подключение библиотеки excel и создание псевдонима "Alias"
using word = Microsoft.Office.Interop.Word; // подключение библиотеки word и создание псевдонима "Alias"

namespace WindowsFormsApplication1
{
    public struct Tema
    {
        public string Name;// ' Название темы
        public string Text; // ' Содержание темы
        public string Rez;// As String ' Результат темы
        public string Comp;// As String ' Компетенции, развиваемые темой
        public string FormZ;// As String ' Формы занятий
        public int N_Sem; // As Integer  ' Номер семестра
    }
    
    struct Plan
    {
        public string AimsDist;
        public string Tasks;
        public string ZnatYmetVladetBefore;
        public string ZnatYmetVladetAfter;
        public string PlanResults;
        public string ContentDist;// массив
        public string LiteraBasic;//массив
        public string LiteraAdditional;//массив
        public string strWord1;
        public string strWord2;
        public List<string> Litera;
        public void CreateLitera()
        {
            Litera = new List<string>();
        }
        public void MyListAdd(string Val)
        {
            Litera.Add(Val); 
        }
    }
    
    public partial class Form1 : Form
    {
        Tema tems;
        public Dis D = new Dis();
        char[] MyChar = { '\f', '\n', '\r', '\t', '\v', '\0', ' ', '2', '3', '.', ')', ';' };
        int CountKFind;  //' счетчик найденных фрагментов, n-сколько надо отсчитать нахождений до нужного
        word.Application WordApp;
        Plan PL;
        public Form1()
        {
            InitializeComponent();
        }
    
        public string SearchText(string wordText1, string wordText2, int nf) // Поиск между двумя фрагментами
        {
            Microsoft.Office.Interop.Word.Range r;//Range
            string st;
            st = "";
            r = WordApp.ActiveDocument.Range();
            bool f;
            f = false;
            int firstOccurence;
            firstOccurence = 0;
            CountKFind = 0;
            r.Find.ClearFormatting(); //Сброс форматирований из предыдущих операций поиска
            r.Find.Text = wordText1 + "*" + wordText2;
            r.Find.Forward = true;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.Format = false;
            r.Find.MatchCase = false;
            r.Find.MatchWholeWord = false;
            r.Find.MatchAllWordForms = false;
            r.Find.MatchSoundsLike = false;
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            while (r.Find.Execute() == true) // Проверка поиска, если нашёл фрагменты, то...
            {
                CountKFind = CountKFind + 1;// то счётчик найденных фрагментоd увеличивается на 1
                if (f) 
                {
                    if (r.Start == firstOccurence) 
                    { }
                    else
                    {
                        firstOccurence = r.Start;
                        f = true;
                    }
                }
                st = WordApp.ActiveDocument.Range(r.Start + wordText1.Length, r.End - wordText2.Length).Text; //убираем кл.
                r.Start = r.Start + wordText1.Length;
                r.End = r.End - wordText2.Length;
                if (CountKFind >= nf) // если нужный по счету фрагмент найден
                {
                    r = WordApp.ActiveDocument.Range(r.Start, r.End);
                    richTextBox1.Text = richTextBox1.Text + r.Text;
                    //MessageBox.Show(CountKFind.ToString());
                }
                //else
                //{
                //    if (CountKFind <= nf)
                //    {
                //        MessageBox.Show("Текст не найден!");
                //        st = ""; // Тут искать ошибку, не выводит текст.
                //    }
                //}
            }
            if (CountKFind == 0)
            {
                if (r.Text != "")
                {
                    if (st != "")
                    {
                        r.Copy();
                    }
                    else //' если текст не найден очистим буфер обмена
                    {
                        Clipboard.Clear();
                    }
                }
                else
                {
                    {
                        Clipboard.Clear();
                    }
                }
            }
            return st;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string Filename; 
            WordApp = new word.Application(); // создаем объект word;
            WordApp.Visible = true; // показывает или скрывает файл word;
            openFileDialog1.ShowDialog();
            Filename = openFileDialog1.FileName;
            WordApp.Documents.Add(Filename);// загружаем в word файл с рабочей книгой 

            //Thread theard = new Thread(SearchText); //второй поток для 
            //theard.Start();
            //Action action = () => { openFileDialog1.ShowDialog(); }; Invoke(action);  // Запуск главного потока 
        }
            
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
        //    int k;
        //    k = Int32.Parse(textBox3.Text);
        //    string findText = textBox2.Text;
        //    SearchText(textBox2.Text, textBox4.Text, k);
        //    int intFound = 0;
        //    WordApp.Selection.Start = 0;
        //    WordApp.Selection.End = 0;
        //    WordApp.Selection.Select();
        //    WordApp.Selection.Find.ClearFormatting();    
        //    if (WordApp.Selection.Find.Execute(findText, Forward:true, MatchWildcards:true, Wrap:word.WdFindWrap.wdFindContinue))
        //    {
        //        MessageBox.Show("Text found.");
        //        int f1;
        //        int f2;
                
        //        f1 = findText.IndexOf("*");
        //        f2 = findText.Length-f1+2;
        //        WordApp.Selection.Start = WordApp.Selection.Start + f1;
        //        WordApp.Selection.End = WordApp.Selection.End - f2;
        //        WordApp.Selection.Select();

        //    }
        //    else
        //    {
        //        MessageBox.Show("The text could not be located.");
        //    }
        //    int p1 = 0;
        //    int p2 = 0;
        //   while(WordApp.Selection.Find.Found)
        //   {
        //       bool prov;
        //       intFound++;
        //  prov = WordApp.Selection.Find.Execute(findText, Forward: true, MatchWildcards: true, Wrap: word.WdFindWrap.wdFindContinue);
        //  int f1;
        //  int f2;
         
        //       if (intFound == 1)
        //       {
        //           p1 = WordApp.Selection.Start;
        //       }
        //       else 
        //       {
        //           p2 = WordApp.Selection.Start;
        //       }
        //        MessageBox.Show("Strings found: " + intFound.ToString());
        //       if (p1 == p2)
        //       {
                   
        //           break;
        //       }
        //       f1 = findText.IndexOf("*");
        //       f2 = findText.Length - f1 + 1;
        //       WordApp.Selection.Start = WordApp.Selection.Start + f1;
        //       WordApp.Selection.End = WordApp.Selection.End - f2;
        //       string vd = WordApp.Selection.Text;
        //       richTextBox1.Text = richTextBox1.Text + vd.Trim(); 

        //   }
           

           
            
        
        }

        private void button3_Click(object sender, EventArgs e)
        { 
            
 
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            //Action action1 = () => { MessageBox.Show("Complete"); }; Invoke(action1); // Запуск главного потока 
            
            SearchText(textBox2.Text, textBox4.Text, CountKFind);
            int N = 0;
            int i = 0;
            int j = 0;
            Microsoft.Office.Interop.Word.Range r;//Range
            Microsoft.Office.Interop.Word.ListParagraphs p;
            PL.CreateLitera();
            string ss;
            ss = "";
            r = WordApp.ActiveDocument.Range();
            p = WordApp.ActiveDocument.ListParagraphs;
            word.Document document = WordApp.ActiveDocument;
            int NnN = document.ListParagraphs.Count;
            ss = SearchText("Основная литература","Перечень",2);
            string str1 = "Основная литература";
            string str2 = "Перечень";
            string gg;
            
            
            r.Find.Text = str1 + "*" + str2;
            r.Find.Forward = true;
            string f = r.Find.Text;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            if (r.Find.Execute(f))// Проверка поиска, если нашёл фрагменты, то...
            {

                gg = WordApp.ActiveDocument.Range(r.Start + str1.Length, r.End - str2.Length).Text; //убираем кл.
                r.Start = r.Start + str1.Length;
                r.End = r.End - str2.Length;
                int mmmmmm = r.ListParagraphs.Count;
                for (int y = 1; y <= r.ListParagraphs.Count;  y++)
                {
                    string dfs = r.ListParagraphs[y].Range.Text;
                    PL.MyListAdd(dfs);
                    richTextBox4.Text = richTextBox4.Text + PL.Litera[y-1];
                }
            }
            

            //for (j = 1; j <= NnN; j++ )
            //{
                
            //    r.ListFormat.ConvertNumbersToText();
            //    string mmmmmm = WordApp.ActiveDocument.ListParagraphs[j].Range.Text;

            //    MessageBox.Show("Абзацев " + NnN + " " + mmmmmm);
            //}
           // word.Range firstRange = document.Paragraphs[1].Range;
           // word.Range secondRange = document.Paragraphs[2].Range;
            

            //string firstString = firstRange.Text;
            //string secondString = secondRange.Text;
           // richTextBox3.Text = richTextBox3.Text + firstString + secondString;
            

            if (ss == "") //' Если цели не попали в оглавление
            {
                ss = SearchText("явля?????", "Учебные задачи дисциплины", 2);
            }

                ss = ss.TrimEnd(MyChar);
                N = ss.IndexOf("явля");
                if (N > 0 && N < ss.Length - 9)
                {
                    D.Cel = ss.Remove(1, N + 9);
                }
                else
                {
                    D.Cel = ss;
                }

            
        
        //' Находим задачи и оставляем все после слова "является" или "являются:"
        ss = SearchText("Учебные задачи дисциплины", "Место дисциплины", 2);
        if (ss == "")// ' Если задачи не попали в оглавление
        {
            ss = SearchText("Учебные задачи дисциплины", "Место дисциплины", 1);
        }
        
        ss = ss.TrimEnd(MyChar);
        N = ss.IndexOf("явля");
        
        if (N > 0 && N < ss.Length - 9) 
    {
        D.Tasks = ss.Remove(1, N + 9);        
    }
        else
        {
            D.Tasks = ss; 
        }
            //Находим знания, умения и владения и оставляем все до знаков препинания и символов перевода, или цифр 2, 3.
            ss = SearchText("Знать:", "Уметь:", 1);
        D.Zn_before = ss.TrimEnd(MyChar);
        ss = SearchText("Уметь:", "Владеть:", 1);
        D.Um_before = ss.TrimEnd(MyChar);
        ss = SearchText("Владеть:", ".", 1);
        D.Vl_before = ss.TrimEnd(MyChar);
        ss = SearchText("Знать:", "Уметь:", 2);
        D.Zn_after = ss.TrimEnd(MyChar);
        ss = SearchText("Уметь:", "Владеть:", 2);
        D.Um_after = ss.TrimEnd(MyChar);
        ss = SearchText("Владеть:", ".", 2);
        D.Vl_after = ss.TrimEnd(MyChar);
//Tema tt = new Tema();
        byte razd = 1;  //'номер раздела
        for (i = 2; i <= WordApp.ActiveDocument.Tables[2].Rows.Count; i++)
        {
            if (WordApp.ActiveDocument.Tables[2].Rows[i].Cells.Count>=5)
            {
                D.tems[i - 2].Name = WordApp.ActiveDocument.Tables[2].Cell(i, 2).Range.Text;
                D.tems[i - 2].Text = WordApp.ActiveDocument.Tables[2].Cell(i, 3).Range.Text;
                //D.tems[i - 2].N_Sem = D.First_Sem + razd - 1 ' попытка определить семестр для данной темы (предполагается каждый раздел в своем семестре)
                //            ' компетенции надо брать из учебного плана, а не из старой программы
                //            ' tt.Comp = .Cell(i, 4).Range.Text
                //            tt.Comp = ""
                //            For j = 0 To D.Nc - 1 ' цикл по компетенциям данной дисциплины из учебного плана
                //                If j <> D.Nc - 1 And D.Nc > 1 Then ' либо через запятую 
                //                    tt.Comp = tt.Comp & D.List_Comp(j).Index & ", "
                //                Else ' либо без знака если компетенция одна или последняя 
                //                    tt.Comp = tt.Comp & D.List_Comp(j).Index
                //   End If
                //Next j
            }
            else 
            {
                if (i != 2) 
                {
                    razd += razd;  //' счетчик разделов срабатывает если их больше одного
                }
            }
              D.tems[i - 2].Rez = WordApp.ActiveDocument.Tables[2].Cell(i, 5).Range.Text;
              D.tems[i - 2].FormZ = WordApp.ActiveDocument.Tables[2].Cell(i, 6).Range.Text;
              richTextBox2.Text = richTextBox2.Text + D.tems[i - 2].Name + D.tems[i - 2].Text + D.tems[i - 2].Rez + D.tems[i - 2].FormZ;
              Clipboard.SetText(richTextBox2.Text + D.tems[i - 2].Name + D.tems[i - 2].Text + D.tems[i - 2].Rez + D.tems[i - 2].FormZ);    
        }
        }

        public void Process()
        {
            Application.Run();
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        } 
   }
}
