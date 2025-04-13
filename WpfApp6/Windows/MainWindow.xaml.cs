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


    }
}
