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

namespace _322_Petrov
{
    /// <summary>
    /// Логика взаимодействия для ChangePassPage.xaml
    /// </summary>
    public partial class ChangePassPage : Page
    {
        public ChangePassPage()
        {
            InitializeComponent();
        }
        public static string GetHash(string password)
        {
            using (var hash = SHA1.Create())
            {
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x => x.ToString("X2")));
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(CurrentPasswordBox.Password) ||
                string.IsNullOrEmpty(NewPasswordBox.Password) ||
                string.IsNullOrEmpty(ConfirmPasswordBox.Password) ||
                string.IsNullOrEmpty(TbLogin.Text))
            {
                MessageBox.Show("Все поля обязательны к заполнению!");
                return;
            }

            string hashedPass = GetHash(CurrentPasswordBox.Password);
            var user = Entities.GetContext().Users
                .FirstOrDefault(u => u.Login == TbLogin.Text && u.Password == hashedPass);

            if (user == null)
            {
                MessageBox.Show("Текущий пароль/Логин неверный!");
                return;
            }

            if (NewPasswordBox.Password.Length < 6)
            {
                MessageBox.Show("Новый пароль слишком короткий, должно быть минимум 6 символов!");
                return;
            }

            bool en = true;
            bool number = false;

            for (int i = 0; i < NewPasswordBox.Password.Length; i++)
            {
                if (NewPasswordBox.Password[i] >= '0' && NewPasswordBox.Password[i] <= '9')
                    number = true;
                else if (!((NewPasswordBox.Password[i] >= 'A' && NewPasswordBox.Password[i] <= 'Z') ||
                          (NewPasswordBox.Password[i] >= 'a' && NewPasswordBox.Password[i] <= 'z')))
                    en = false;
            }

            if (!en)
            {
                MessageBox.Show("Используйте только английскую раскладку для нового пароля!");
                return;
            }

            if (!number)
            {
                MessageBox.Show("Добавьте хотя бы одну цифру в новый пароль!");
                return;
            }

            if (NewPasswordBox.Password != ConfirmPasswordBox.Password)
            {
                MessageBox.Show("Новые пароли не совпадают!");
                return;
            }

            try
            {
                user.Password = GetHash(NewPasswordBox.Password);
                Entities.GetContext().SaveChanges();
                MessageBox.Show("Пароль успешно изменен!");
                NavigationService?.Navigate(new Pages.AuthPage());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при изменении пароля: {ex.Message}");
            }
        }
    }
}
