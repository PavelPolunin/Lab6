using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using WorkWithWord.HelperClasses;
using WpfApp6.ModelClasses;
using Microsoft.Office.Interop.Word;

namespace WpfApp6
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ModelEF model;

        private List<Users> users;

        private List<Auto> autos;

        public MainWindow()
        {
            InitializeComponent();

            model = new ModelEF();

            users = new List<Users>();
            autos = new List<Auto>();
        }

        private void ComboLoadData()
        {
            comboBoxUsers.Items.Clear();

            users = model.Users.ToList();

            foreach (var item in users) 
                comboBoxUsers.Items.Add($"{item.FullName} {item.PSeria} {item.PNumber}");

            comboBoxUsers.SelectedIndex = 0;

            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();
            comboBoxAutos.Items.Clear();

            foreach (var item in autos) 
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN} ");

            comboBoxAutos.SelectedIndex = 0;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ComboLoadData();
        }

        private void comboBoxUsers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();

            comboBoxAutos.Items.Clear();

            foreach(var item in autos) 
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN} ");

            comboBoxAutos.SelectedIndex = 0;
        }

        private void SaveDocument_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.Description = "Выберите место сохранения";


            if (System.Windows.Forms.DialogResult.OK == fbd.ShowDialog())
            {
                Users activeUser = users[comboBoxUsers.SelectedIndex];
                Auto activeAuto = activeUser.Auto.ToList()[comboBoxAutos.SelectedIndex];

                CreateDocument(
                    $@"{fbd.SelectedPath}\Купля-Продажа-Автомобиля-{activeUser.FullName}.docx", 
                    activeUser, 
                    activeAuto);
                System.Windows.MessageBox.Show("Файл сохранён");
            }
        }

        private void CreateDocument(string directorypath, Users users, Auto auto)
        {
            var today = DateTime.Now.ToShortDateString();
            WordHelper word = new WordHelper("ContractSale.docx");
            var items = new Dictionary<string, string>
            {
                {"<Today>", today },
                {"<FullName>", users.FullName }, // ФИО
                {"<Date0fBirth>", users.DateOfBirth.Value.ToShortDateString() }, // Дата рождения
                {"<Adress>", users.Adress },
                {"<PSeria>", users.PSeria.ToString() }, // Серия паспорта
                {"<PNumber>", users.PNumber.ToString() }, // Номер паспорта
                {"<PVidan>", users.PVidan },
                {"<ModelV>", auto.Model },
                {"<CategoryV>", auto.Category },
                {"<TypeV>", auto.TypeV },
                {"<VIN>", auto.VIN },
                {"<RegistrationMark>", auto.RegistrationMark }, // Регистрационный знак
                {"<YearV>", auto.YearOfRelease.Value.Year.ToString() }, // Год выпуска
                {"<EngineV>", auto.EngineNumber }, // Номер двигателя
                {"<ChassisV>", auto.Chassis },
                {"<BodyworkV>", auto.Bodywork },
                {"<ColorV>", auto.Color },
                {"<SeriaPV>", auto.SeriaPasport }, // Серия ПТС
                {"<NumberPV>", auto.NumbePasport }, // Номер ПТС
                {"<VidanPV>", auto.VidanPasport } // Кем выдан ПТС
            };
            word.Process(items, directorypath);
        }
    }
}
