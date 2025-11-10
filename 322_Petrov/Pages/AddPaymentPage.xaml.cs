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

namespace _322_Petrov.Pages
{
    /// <summary>
    /// Логика взаимодействия для AddPaymentPage.xaml
    /// </summary>
    public partial class AddPaymentPage : Page
    {
        private Payment _currentPayment = new Payment();

        public AddPaymentPage(Payment selectedPayment)
        {
            InitializeComponent();
            LoadComboBoxData();

            if (selectedPayment != null)
                _currentPayment = selectedPayment;

            DataContext = _currentPayment;
        }

        private void LoadComboBoxData()
        {
            try
            {
                var context = new Entities();
                CBCategory.ItemsSource = context.Categories.ToList();
                CBUser.ItemsSource = context.Users.ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки данных: {ex.Message}");
            }
        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            StringBuilder errors = new StringBuilder();

            // Проверка обязательных полей
            if (string.IsNullOrWhiteSpace(_currentPayment.Name))
                errors.AppendLine("Укажите название платежа!");
            if (_currentPayment.Date == null || _currentPayment.Date == DateTime.MinValue)
                errors.AppendLine("Укажите корректную дату!");
            if (_currentPayment.Num <= 0)
                errors.AppendLine("Укажите корректное количество!");
            if (_currentPayment.Price <= 0)
                errors.AppendLine("Укажите корректную цену!");
            if (_currentPayment.UserID == 0)
                errors.AppendLine("Укажите клиента!");
            if (_currentPayment.CategoryID == 0)
                errors.AppendLine("Укажите категорию!");

            if (errors.Length > 0)
            {
                MessageBox.Show(errors.ToString());
                return;
            }

            try
            {
                var context = new Entities();

                if (_currentPayment.ID == 0)
                {
                    context.Payments.Add(_currentPayment);
                }
                else
                {
                    var existingPayment = context.Payments.Find(_currentPayment.ID);
                    if (existingPayment != null)
                    {
                        context.Entry(existingPayment).CurrentValues.SetValues(_currentPayment);
                    }
                }

                context.SaveChanges();
                MessageBox.Show("Данные успешно сохранены!");

                // Возврат на предыдущую страницу
                NavigationService?.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения: {ex.Message}");
            }
        }

        private void ButtonClean_Click(object sender, RoutedEventArgs e)
        {
            // Создаем новый объект вместо очистки полей вручную
            _currentPayment = new Payment();
            DataContext = _currentPayment;
            CBCategory.SelectedIndex = -1;
            CBUser.SelectedIndex = -1;
        }
    }
}
