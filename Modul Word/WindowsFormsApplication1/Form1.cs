﻿using System;
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

    public struct Discipline
    {
        public string Index;// 'Индекс (номер дисциплины в плане)
        public string Name;// 'Наименование
        public string Exam;// 'Экзамены
        public string Zach;// 'Зачеты
        public string Zach_E;// 'Зачеты с оценкой
        public string Section;// 'Раздел плана
        public string Curs_R;// ' Курсовые работы
        public string Cafedra;// 'Закрепленная кафедра
        public byte First_Sem;// 'Первый семестр изучения дисциплины
        public byte Last_Sem;//'Последний семестр изучения дисциплины
        public string List_Comp;// 'Список компетенций
    }
    
    public partial class Form1 : Form
    {
        Tema tems;
        Discipline dis;
        public Dis D = new Dis(); /*Класс*/
        char[] MyChar = { '\f', '\n', '\r', '\t', '\v', '\0', ' ', '2', '3', '.', ')', ';' };
        int CountKFind;  //' счетчик найденных фрагментов, n-сколько надо отсчитать нахождений до нужного
        word.Application WordApp;
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
                    //richTextBox1.Text = richTextBox1.Text + r.Text;
                }
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
        } // Открытие word документа
            
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
           

           
            
        
        } // Первый вариант поиска 

        private void button3_Click(object sender, EventArgs e)
        { 
            
 
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e) // Основная кнопка поиска
        {
            
            //Action action1 = () => { MessageBox.Show("Complete"); }; Invoke(action1); // Запуск главного потока 
            
            SearchText(textBox2.Text, textBox4.Text, CountKFind);
            int N = 0;
            int i = 0;
            int j = 0;
            Microsoft.Office.Interop.Word.Range r;//Range
            Microsoft.Office.Interop.Word.ListParagraphs p;
            D.CreateLitera();
            string ss;
            ss = "";
            r = WordApp.ActiveDocument.Range();
            p = WordApp.ActiveDocument.ListParagraphs;
            word.Document document = WordApp.ActiveDocument;
            int NnN = document.ListParagraphs.Count;
            //Поиск литературы
            string str1 = "Основная литература";
            string str2 = "Дополнительная литература";
            string str3 = "Перечень";
            string gg1; string gg2;

            // Поиск основной литературы
            r.Find.Text = str1 + "*" + str2;
            r.Find.Forward = true;
            string f1 = r.Find.Text;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            
            if (r.Find.Execute(f1))// Проверка поиска, если нашёл фрагменты, то...
            {
                

                gg1 = WordApp.ActiveDocument.Range(r.Start + str1.Length, r.End - str2.Length).Text; //убираем кл.
                r.Start = r.Start + str1.Length;
                r.End = r.End - str2.Length;
                int m21 = r.ListParagraphs.Count;
                for (int y = 1; y <= r.ListParagraphs.Count;  y++)
                {
                    //MessageBox.Show(""+ m21);
                    string dfs = r.ListParagraphs[y].Range.Text;
                    D.MyListAdd(dfs, false);
                    richTextBox4.Text = richTextBox4.Text + D.LiteraBasic[y-1];
                }
            }
            // поиск дополнительной литературы
            r.Find.Text = str2 + "*" + str3;
            r.Find.Forward = true;
            string f2 = r.Find.Text;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            if (r.Find.Execute(f2))// Проверка поиска, если нашёл фрагменты, то...
            {
                

                gg2 = WordApp.ActiveDocument.Range(r.Start + str2.Length, r.End - str3.Length).Text; //убираем кл.
                r.Start = r.Start + str2.Length;
                r.End = r.End - str3.Length;
                int m12 = r.ListParagraphs.Count;
                for (int x = 1; x <= r.ListParagraphs.Count; x++)
                {
                    //MessageBox.Show(""+ m12);
                    string dsf = r.ListParagraphs[x].Range.Text;
                    D.MyListAdd(dsf, true);
                    richTextBox5.Text = richTextBox5.Text + D.LiteraAdditional[x-1];
                }
            }


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
        int CountTems = 0;
        for (i = 2; i <= WordApp.ActiveDocument.Tables[2].Rows.Count; i++)
        {
            if (WordApp.ActiveDocument.Tables[2].Rows[i].Cells.Count >= 5)
            {
                D.tems[i - 2].Name = WordApp.ActiveDocument.Tables[2].Cell(i, 2).Range.Text;
                D.tems[i - 2].Text = WordApp.ActiveDocument.Tables[2].Cell(i, 3).Range.Text;
                D.tems[i - 2].Rez = WordApp.ActiveDocument.Tables[2].Cell(i, 5).Range.Text;
                D.tems[i - 2].FormZ = WordApp.ActiveDocument.Tables[2].Cell(i, 6).Range.Text;
                CountTems++;
                richTextBox2.Text = richTextBox2.Text + D.tems[i - 2].Name + D.tems[i - 2].Text + D.tems[i - 2].Rez + D.tems[i - 2].FormZ;
                Clipboard.SetText(richTextBox2.Text + D.tems[i - 2].Name + D.tems[i - 2].Text + D.tems[i - 2].Rez + D.tems[i - 2].FormZ);
                
                
            }
            else 
            {
                if (i != 2) 
                {
                    razd += razd;  //' счетчик разделов срабатывает если их больше одного
                }
            }     
        }
        Clipboard.Clear();

        ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 2);
        Clipboard.SetText(ss);
        
            int n1, n2, n3, n4;
            n1 = ss.IndexOf("Тема");
            n2 = ss.IndexOf("Литература");
            n3 = ss.IndexOf("Вопросы для");
            n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
            if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10)  //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
            {
                richTextBox3.Text = "";
                richTextBox3.Paste();
                if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) // ' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                {
                    richTextBox3.Text = "";
                    richTextBox3.Paste();
                    if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                    {
                        richTextBox3.Text = "";
                        richTextBox3.Paste();
                        if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                        {
                            richTextBox3.Text = "";
                            richTextBox3.Paste();
                            if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                            {
                                richTextBox3.Text = "";
                                richTextBox3.Paste();
                                if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > n2) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                                {
                                    richTextBox3.Text = "";
                                    richTextBox3.Paste();
                                    if ((n1 > 0 && n1 < 100) &&(n2 > n1 && n2 < 300) && n3 > n2)  //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                                    {
                                        richTextBox3.Text = "";
                                        richTextBox3.Paste();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Перечень УМО не найден");
                                        richTextBox3.Text = "";
                                    }
                                }
                                else
                                {
                                    ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ", 1);
                                    n1 = ss.IndexOf("Тема");
                                    n2 = ss.IndexOf("Литература");
                                    n3 = ss.IndexOf("Вопросы для");
                                }
                            }
                            else //' Это для РП образца 2015г.
                            {
                                ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ", 2);
                                n1 = ss.IndexOf("Тема");
                                n2 = ss.IndexOf("Литература");
                                n3 = ss.IndexOf("Вопросы для");
                            }
                        }
                        else
                        {
                            ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Рекомендуемые обучающие", 1);
                            n1 = ss.IndexOf("Тема");
                            n2 = ss.IndexOf("Литература");
                            n3 = ss.IndexOf("Вопросы для");
                            n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
                        }
                    }
                    else
                    {
                        ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Рекомендуемые обучающие", 2);
                        n1 = ss.IndexOf("Тема");
                        n2 = ss.IndexOf("Литература");
                        n3 = ss.IndexOf("Вопросы для");
                        n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
                    }
                }
                else //' это если в конце файла есть еще раз этот раздел, то надо искать третье вхождение
                {
                    ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 3);
                    n1 = ss.IndexOf("Тема");
                    n2 = ss.IndexOf("Литература");
                    n3 = ss.IndexOf("Вопросы для");
                    n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
                }
            }
            else //' это если в содержании нет этого раздела, а в тексте есть
            {
                ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 1);
                n1 = ss.IndexOf("Тема");
                n2 = ss.IndexOf("Литература");
                n3 = ss.IndexOf("Вопросы для");
                n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
            }


            Clipboard.Clear();
            int k, k1; // ' метки для найденных символов
            k = richTextBox2.Find("Тема 1");
            richTextBox2.SelectAll();
            if (k > 0) 
            {
                richTextBox2.SelectedText.Remove(0, k - 1);
            }
            ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ ДЛЯ ОБУЧАЮЩИХСЯ", 2);
            Clipboard.SetText(ss);
        if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
          {
                richTextBox3.Text = "";
                richTextBox3.Paste();
            if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
            {
                richTextBox3.Text = "";
                richTextBox3.Paste();
            
                if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                {
                    richTextBox3.Text = "";
                    richTextBox3.Paste();
                    if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                    {
                        richTextBox3.Text = "";
                        richTextBox3.Paste();
                        if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                        {
                                richTextBox3.Text = "";
                                richTextBox3.Paste();
                            if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
                            {
                                    richTextBox3.Text = "";
                                    richTextBox3.Paste();
                                if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
                                {
                                        richTextBox3.Text = "";
                                        richTextBox3.Paste();
                                    if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                                    {
                                            richTextBox3.Text = "";
                                            richTextBox3.Paste();
                                    }
                                        else
                                        {
                                            MessageBox.Show("Перечень Заданий не найден!");
                                            richTextBox3.Text = "";
                                        }
                                }
                                else
                                        {
                                        ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "Тематический план", 1);
                                        }
                            }
                            else
                                    {
                                    ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "Тематический план", 2);
                                    }
                        }
                        else
                                {
                                 ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "ТЕМАТИЧЕСКИЙ ПЛАН", 1);
                                }
                    }
                    else //' это для РПД образца 2015г.
                            {
                            ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "ТЕМАТИЧЕСКИЙ ПЛАН", 2);
                            }
                }
                else
                        {
                        ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "МЕТОДИЧЕСКИЕ УКАЗАНИЯ", 1);
                        }
            }
                else
                    {
                    ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "МЕТОДИЧЕСКИЕ УКАЗАНИЯ", 2);
                    }
        }
            else
                {
                ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ ДЛЯ ОБУЧАЮЩИХСЯ", 1);
                }







        Clipboard.Clear();
        //Поиск вопросов к экзамену/зачёту
        string exstr1 = "Вопросы к";
        string exstr2 = "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ";
        string exgg1; 

        // Поиск 
        r.Find.Text = exstr1 + "*" + exstr2;
        r.Find.Forward = true;
        string exf1 = r.Find.Text;
        r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
        r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ

        if (r.Find.Execute(exf1))// Проверка поиска, если нашёл фрагменты, то...
        {


            exgg1 = WordApp.ActiveDocument.Range(r.Start + exstr1.Length, r.End - exstr2.Length).Text; //убираем кл.
            r.Start = r.Start + exstr1.Length;
            r.End = r.End - exstr2.Length;
            int exm21 = r.ListParagraphs.Count;
            for (int y = 1; y <= r.ListParagraphs.Count; y++)
            {
                //MessageBox.Show("" + exm21);
                string dfs = r.ListParagraphs[y].Range.Text;
                D.MyForExamAdd(dfs);
                richTextBox1.Text = richTextBox1.Text + D.ForExam[y-1];
            }
        }



            //' Теперь надо выделить вопросы к экзамену и записать их в массив вопросов
            //int k0;
            //string ss1;
            //if (richTextBox3.Text != "")
            //{
            //    k0 = richTextBox3.Find("Вопросы");
            //    if (k0 > 0)  //' манипуляции с k0, k, k1 для повышения надежности поиска вопросов
            //    {
            //        int m; //'сколько знаков вырезать с номером вопроса
            //        for (i = 1; i <= 100; i++)
            //        {
            //            if (i < 10)
            //            {
            //                m = 3;
            //            }
            //            else
            //            {
            //                m = 4;
            //            }
            //            //}

            //            ss = i + ".";
            //            k = richTextBox3.Find(ss, k0, RichTextBoxFinds.NoHighlight);
            //            ss = (i + 1) + ".";
            //            k1 = richTextBox3.Find(ss, k, RichTextBoxFinds.NoHighlight);
            //            if (k1 > k + m) // ' если очередной вопрос найден то выделим его и через буфер обмена перенесем в массив
            //            {
            //                richTextBox3.Select(k + m, k1 - k - m);
            //                richTextBox3.Copy();
            //                ss1 = Clipboard.GetText();
            //                D.MyForExamAdd(ss1);
            //            }
            //            else
            //            {
            //                //Exit For
            //            }

            //            k0 = k;
            //            //Next i
                        
            //        }
            //    }
            //}
            //    /*
            //        WordApp.ActiveDocument.Close();
            //        WordApp.Quit();
            //     */
            //    //try
            //    //{
            //    //}
            //    //catch //' попытка отстроится от ошибок и файл все равно закрыть
            //    //{
            //    //    WordApp.ActiveDocument.Close();
            //    //    WordApp.Quit();
            //    //}
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
       /* public string SearchExam()
        {
            //Поиск вопросов к экзамену/зачёту
            string str1 = "Вопросы к";
            string str2 = "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ";
            string gg1; string gg2;

            // Поиск 
            r.Find.Text = str1 + "*" + str2;
            r.Find.Forward = true;
            string f1 = r.Find.Text;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ

            if (r.Find.Execute(f1))// Проверка поиска, если нашёл фрагменты, то...
            {


                gg1 = WordApp.ActiveDocument.Range(r.Start + str1.Length, r.End - str2.Length).Text; //убираем кл.
                r.Start = r.Start + str1.Length;
                r.End = r.End - str2.Length;
                int m21 = r.ListParagraphs.Count;
                for (int y = 1; y <= r.ListParagraphs.Count; y++)
                {
                    MessageBox.Show(""+ m21);
                    string dfs = r.ListParagraphs[y].Range.Text;
                    D.MyListAdd(dfs, false);
                    richTextBox4.Text = richTextBox4.Text + D.ForExam[y - 1];
                }
            }
            // поиск дополнительной литературы
            r.Find.Text = str2 + "*" + str3;
            r.Find.Forward = true;
            string f2 = r.Find.Text;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            if (r.Find.Execute(f2))// Проверка поиска, если нашёл фрагменты, то...
            {


                gg2 = WordApp.ActiveDocument.Range(r.Start + str2.Length, r.End - str3.Length).Text; //убираем кл.
                r.Start = r.Start + str2.Length;
                r.End = r.End - str3.Length;
                int m12 = r.ListParagraphs.Count;
                for (int x = 1; x <= r.ListParagraphs.Count; x++)
                {
                    //MessageBox.Show(""+ m12);
                    string dsf = r.ListParagraphs[x].Range.Text;
                    D.MyListAdd(dsf, true);
                    richTextBox5.Text = richTextBox5.Text + D.LiteraAdditional[x - 1];
                }
            }
        }*/
   }
}

