using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
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
    /// Логика взаимодействия для RegPage.xaml
    /// </summary>
    public partial class RegPage : Page
    {
        public RegPage()
        {
            InitializeComponent();
            comboBxRole.SelectedIndex = 0;
        }

        public static string GetHash(String password)
        {
            using (var hash = SHA1.Create())
            {
                return
                string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x =>
                x.ToString("X2")));
            }
        }

        private void lblLogHitn_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            txtbxLog.Focus();
        }

        private void lblPassHitn_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            passBxFrst.Focus();
        }

        private void lblPassSecHitn_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            passBxScnd.Focus();
        }

        private void lblFioHitn_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            txtbxFIO.Focus();
        }

        // Обработчики изменения текста
        private void txtbxLog_TextChanged(object sender, TextChangedEventArgs e)
        {
            lblLogHitn.Visibility = string.IsNullOrEmpty(txtbxLog.Text) ? Visibility.Visible : Visibility.Hidden;
        }

        private void txtbxFIO_TextChanged(object sender, TextChangedEventArgs e)
        {
            lblFioHitn.Visibility = string.IsNullOrEmpty(txtbxFIO.Text) ? Visibility.Visible : Visibility.Hidden;
        }

        private void passBxFrst_PasswordChanged(object sender, RoutedEventArgs e)
        {
            lblPassHitn.Visibility = string.IsNullOrEmpty(passBxFrst.Password) ? Visibility.Visible : Visibility.Hidden;
        }

        private void passBxScnd_PasswordChanged(object sender, RoutedEventArgs e)
        {
            lblPassSecHitn.Visibility = string.IsNullOrEmpty(passBxScnd.Password) ? Visibility.Visible : Visibility.Hidden;
        }

        // Обработчики получения фокуса
        private void txtbxLog_GotFocus(object sender, RoutedEventArgs e)
        {
            lblLogHitn.Visibility = Visibility.Hidden;
        }

        private void txtbxLog_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtbxLog.Text))
            {
                lblLogHitn.Visibility = Visibility.Visible;
            }
        }

        private void txtbxFIO_GotFocus(object sender, RoutedEventArgs e)
        {
            lblFioHitn.Visibility = Visibility.Hidden;
        }

        private void txtbxFIO_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtbxFIO.Text))
            {
                lblFioHitn.Visibility = Visibility.Visible;
            }
        }

        private void passBxFrst_GotFocus(object sender, RoutedEventArgs e)
        {
            lblPassHitn.Visibility = Visibility.Hidden;
        }

        private void passBxFrst_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(passBxFrst.Password))
            {
                lblPassHitn.Visibility = Visibility.Visible;
            }
        }

        private void passBxScnd_GotFocus(object sender, RoutedEventArgs e)
        {
            lblPassSecHitn.Visibility = Visibility.Hidden;
        }

        private void passBxScnd_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(passBxScnd.Password))
            {
                lblPassSecHitn.Visibility = Visibility.Visible;
            }
        }

        private void comboBxRole_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void regButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtbxLog.Text) ||
                string.IsNullOrEmpty(txtbxFIO.Text) ||
                string.IsNullOrEmpty(passBxFrst.Password) ||
                string.IsNullOrEmpty(passBxScnd.Password))
            {
                MessageBox.Show("Заполните все поля!");
                return;
            }

            Entities db = new Entities();
            var user = db.Users
                .AsNoTracking()
                .FirstOrDefault(u => u.Login == txtbxLog.Text);

            if (user != null)
            {
                MessageBox.Show("Пользователь с таким логином уже существует!");
                return;
            }

            if (passBxFrst.Password.Length < 6)
            {
                MessageBox.Show("Пароль слишком короткий, должно быть минимум 6 символов!");
                return;
            }

            bool en = true;
            bool number = false;

            for (int i = 0; i < passBxFrst.Password.Length; i++)
            {
                if (passBxFrst.Password[i] >= '0' && passBxFrst.Password[i] <= '9')
                    number = true;
                else if (!((passBxFrst.Password[i] >= 'A' && passBxFrst.Password[i] <= 'Z') ||
                          (passBxFrst.Password[i] >= 'a' && passBxFrst.Password[i] <= 'z')))
                    en = false;
            }

            if (!en)
            {
                MessageBox.Show("Используйте только английскую раскладку!");
                return;
            }

            if (!number)
            {
                MessageBox.Show("Добавьте хотя бы одну цифру!");
                return;
            }

            if (passBxFrst.Password != passBxScnd.Password)
            {
                MessageBox.Show("Пароли не совпадают!");
                return;
            }

            try
            {
                User userObject = new User
                {
                    FIO = txtbxFIO.Text,
                    Login = txtbxLog.Text,
                    Password = GetHash(passBxFrst.Password),
                    Role = comboBxRole.Text
                };

                db.Users.Add(userObject);
                db.SaveChanges();

                MessageBox.Show("Пользователь успешно зарегистрирован!");

                txtbxLog.Clear();
                passBxFrst.Clear();
                passBxScnd.Clear();
                comboBxRole.SelectedIndex = 0;
                txtbxFIO.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при регистрации: {ex.Message}");
            }
        }
    }
}

