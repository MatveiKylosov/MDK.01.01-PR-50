using Word_Kylosov.Models;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace Word_Kylosov.Context
{
    public class OwnerContext : Owner
    {
        public OwnerContext(string FirstName, string LastName, string SurName, int NumberRoom) : base(FirstName, LastName, SurName, NumberRoom){}

        public static List<OwnerContext> AllOwners()
        {
            List<OwnerContext> allOwenrs = new List<OwnerContext>();
            allOwenrs.Add(new OwnerContext("Test1", "Test2", "Test3", 1));
            allOwenrs.Add(new OwnerContext("Test4", "Test5", "Test6", 2));
            allOwenrs.Add(new OwnerContext("Test7", "Test8", "Test9", 3));
            allOwenrs.Add(new OwnerContext("Test10", "Test11", "Test12", 4));
            allOwenrs.Add(new OwnerContext("Test13", "Test14", "Test15", 5));
            allOwenrs.Add(new OwnerContext("Test16", "Test17", "Test18", 6));
            allOwenrs.Add(new OwnerContext("Test19", "Test20", "Test21", 7));
            allOwenrs.Add(new OwnerContext("Test22", "Test23", "Test24", 8));
            allOwenrs.Add(new OwnerContext("Test25", "Test26", "Test27", 9));
            allOwenrs.Add(new OwnerContext("Test28", "Test29", "Test30", 10));
            allOwenrs.Add(new OwnerContext("Test31", "Test32", "Test33", 11));
            allOwenrs.Add(new OwnerContext("Test34", "Test35", "Test36", 12));
            allOwenrs.Add(new OwnerContext("Test37", "Test38", "Test39", 13));
            allOwenrs.Add(new OwnerContext("Test40", "Test41", "Test42", 14));
            allOwenrs.Add(new OwnerContext("Test43", "Test44", "Test45", 15));
            allOwenrs.Add(new OwnerContext("Test46", "Test47", "Test48", 16));
            allOwenrs.Add(new OwnerContext("Test49", "Test50", "Test51", 17));
            allOwenrs.Add(new OwnerContext("Test52", "Test53", "Test54", 18));
            allOwenrs.Add(new OwnerContext("Test55", "Test56", "Test57", 19));

            return allOwenrs;
        }

        public static void Report(string fileName)
        {
            // Создаём приложение
            Word.Application app = new Word.Application();
            // Создаём документ
            Word.Document doc = app.Documents.Add();

            // Создаём заголовок
            Word.Paragraph paraHeader = doc.Paragraphs.Add();
            // Указываем шрифт для заголовка
            paraHeader.Range.Font.Size = 16;
            // Задаём текст для заголовка
            paraHeader.Range.Text = "Список жильцов дома";
            // Указываем положение на странице
            paraHeader.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            // Убираем отступ
            paraHeader.Range.ParagraphFormat.SpaceAfter = 0;
            // Убираем жирность
            paraHeader.Range.Font.Bold = 1;
            // Добавляем на документ
            paraHeader.Range.InsertParagraphAfter();

            // Создаём подзаголовок
            Word.Paragraph paraAddress = doc.Paragraphs.Add();
            // Указываем шрифт
            paraAddress.Range.Font.Size = 14;
            // Задаём текст
            paraAddress.Range.Text = "по адресу: г. Пермь, ул. Луначарского, д. 24";
            // Убираем отступ
            paraAddress.Range.ParagraphFormat.SpaceAfter = 20;
            // Убираем жирность
            paraAddress.Range.Font.Bold = 0;
            // Добавляем на документ
            paraAddress.Range.InsertParagraphAfter();

            // Создаём заголовок
            Word.Paragraph paraCount = doc.Paragraphs.Add();
            // Указываем шрифт
            paraCount.Range.Font.Size = 14;
            // Задаём текст
            paraCount.Range.Text = $"Всего жильцов: {AllOwners().Count}";
            // Указываем положение на странице
            paraCount.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            // Убираем отступ
            paraCount.Range.ParagraphFormat.SpaceAfter = 0;
            // Добавляем на документ
            paraCount.Range.InsertParagraphAfter();

            // Создаём таблицу
            Word.Paragraph tableParagraph = doc.Paragraphs.Add();
            // Добавляем на документ
            Word.Table paymentsTable = doc.Tables.Add(tableParagraph.Range, AllOwners().Count + 1, 4);
            // Указываем границы таблицы
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            // Указываем положение таблицы
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            // Создаём заголовки в таблице
            Cell("№", paymentsTable.Cell(1, 1).Range);
            Cell("Фамилия", paymentsTable.Cell(1, 2).Range);
            Cell("Имя", paymentsTable.Cell(1, 3).Range);
            Cell("Отчество", paymentsTable.Cell(1, 4).Range);

            // Перебираем жильцов
            for (int i = 0; i < AllOwners().Count; i++)
            {
                var owner = AllOwners()[i];

                Cell((i + 1).ToString(), paymentsTable.Cell(2 + i, 1).Range);
                Cell(owner.LastName, paymentsTable.Cell(2 + i, 2).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.FirstName, paymentsTable.Cell(2 + i, 3).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.SurName, paymentsTable.Cell(2 + i, 4).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
            }

            // Сохраняем документ
            doc.SaveAs2(fileName);
            // Закрываем документ
            doc.Close();
            // Закрываем приложение
            app.Quit();
        }

        public static void Cell(string Text, Word.Range Cell, Word.WdParagraphAlignment Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter)
        {
            Cell.Text = Text;
            Cell.ParagraphFormat.Alignment = Alignment;
        }
    }
}
