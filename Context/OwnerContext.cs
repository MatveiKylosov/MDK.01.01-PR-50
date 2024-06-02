using Word_Kylosov.Models;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Windows;

namespace Word_Kylosov.Context
{
    public class OwnerContext : Owner
    {
        public OwnerContext(string FirstName, string LastName, string SurName, int NumberRoom, string PhotoPath) : base(FirstName, LastName, SurName, NumberRoom, PhotoPath) {}

        public static List<OwnerContext> AllOwners()
        {
            List<OwnerContext> allOwenrs = new List<OwnerContext>();
            allOwenrs.Add(new OwnerContext("Test1", "Test2", "Test3", 1, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test4", "Test5", "Test6", 2, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test7", "Test8", "Test9", 3, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test10", "Test11", "Test12", 3, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test13", "Test14", "Test15", 4, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test16", "Test17", "Test18", 5, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test19", "Test20", "Test21", 6, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test22", "Test23", "Test24", 6, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test25", "Test26", "Test27", 7, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test28", "Test29", "Test30", 7, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test31", "Test32", "Test33", 8, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test34", "Test35", "Test36", 9, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test37", "Test38", "Test39", 10, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test40", "Test41", "Test42", 11, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test43", "Test44", "Test45", 12, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test46", "Test47", "Test48", 13, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test49", "Test50", "Test51", 14, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test52", "Test53", "Test54", 15, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test55", "Test56", "Test57", 16, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test58", "Test59", "Test60", 16, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test61", "Test62", "Test63", 17, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test64", "Test65", "Test66", 17, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test67", "Test68", "Test69", 18, "C:\\Users\\matve\\Pictures\\testt.png"));
            allOwenrs.Add(new OwnerContext("Test70", "Test71", "Test72", 1, "C:\\Users\\matve\\Pictures\\testt.png"));
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
            Word.Table paymentsTable = doc.Tables.Add(tableParagraph.Range, AllOwners().Count + 1, 6);
            // Указываем границы таблицы
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            // Указываем положение таблицы
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            // Создаём заголовки в таблице
            Cell("№", paymentsTable.Cell(1, 1).Range);
            Cell("Номер квартиры", paymentsTable.Cell(1, 2).Range);
            Cell("Фамилия", paymentsTable.Cell(1, 3).Range);
            Cell("Имя", paymentsTable.Cell(1, 4).Range);
            Cell("Отчество", paymentsTable.Cell(1, 5).Range);
            Cell("Фото", paymentsTable.Cell(1, 6).Range);

            var allOwners = AllOwners();
            allOwners.Sort((x, y) => x.NumberRoom.CompareTo(y.NumberRoom));

            int LastRoom = -1;

            // Перебираем жильцов
            for (int i = 0; i < AllOwners().Count; i++)
            {
                var owner = allOwners[i];

                Cell((i + 1).ToString(), paymentsTable.Cell(2 + i, 1).Range);
                if(owner.NumberRoom != LastRoom)
                {
                    Cell(owner.NumberRoom.ToString(), paymentsTable.Cell(2 + i, 2).Range);
                    LastRoom = owner.NumberRoom;
                }
                else
                {
                    paymentsTable.Cell(1 + i, 2).Merge(paymentsTable.Cell(2 + i, 2));
                    LastRoom = owner.NumberRoom;
                }
                
                Cell(owner.LastName, paymentsTable.Cell(2 + i, 3).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.FirstName, paymentsTable.Cell(2 + i, 4).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                Cell(owner.SurName, paymentsTable.Cell(2 + i, 5).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);

                // Добавляем фото
                string photoPath = owner.PhotoPath;
                if (!string.IsNullOrEmpty(photoPath))
                {
                    var cellRange = paymentsTable.Cell(2 + i, 6).Range;
                    cellRange.InlineShapes.AddPicture(photoPath);
                }
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
