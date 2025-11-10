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
    /// Логика взаимодействия для AuthPage.xaml
    /// </summary>
    public partial class AuthPage : Page
    {
        private int failedAttempts = 0;
        private User currentUser;
        public AuthPage()
        {
            InitializeComponent();
        }

        // Метод для хэширования пароля
        public static string GetHash(string password)
        {
            using (var hash = SHA1.Create())
            {
                return string.Concat(hash.ComputeHash(Encoding.UTF8.GetBytes(password)).Select(x => x.ToString("X2")));
            }
        }

        // Обработчик кнопки входа
        private void ButtonEnter_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextBoxLogin.Text) || string.IsNullOrEmpty(PasswordBox.Password))
            {
                MessageBox.Show("Введите логин или пароль");
                return;
            }

            string hashedPassword = GetHash(PasswordBox.Password);

            using (var db = new Entities())
            {
                var user = db.Users
                .AsNoTracking()
                .FirstOrDefault(u => u.Login == TextBoxLogin.Text && u.Password ==
                hashedPassword);

                if (user == null)
                {
                    MessageBox.Show("Пользователь с такими данными не найден!");
                    failedAttempts++;

                    if (failedAttempts >= 3)
                    {
                        if (captcha.Visibility != Visibility.Visible)
                        {
                            CaptchaSwitch();
                        }
                        CaptchaChange();
                    }
                    return;
                }
                else
                {
                    MessageBox.Show("Пользователь успешно найден!");
                    currentUser = user;

                    switch (user.Role)
                    {
                        case "User":
                            NavigationService?.Navigate(new UserPage());
                            break;
                        case "Admin":
                            NavigationService?.Navigate(new AdminPage());
                            break;
                        default:
                            MessageBox.Show("Неизвестная роль пользователя");
                            break;
                    }
                }
            }
        }

        public void CaptchaSwitch()
        {
            switch (captcha.Visibility)
            {
                case Visibility.Visible:
                    TextBoxLogin.Clear();
                    PasswordBox.Clear();

                    captcha.Visibility = Visibility.Hidden;
                    captchaInput.Visibility = Visibility.Hidden;
                    labelCaptcha.Visibility = Visibility.Hidden;
                    submitCaptcha.Visibility = Visibility.Hidden;

                    labelLogin.Visibility = Visibility.Visible;
                    labelPass.Visibility = Visibility.Visible;
                    TextBoxLogin.Visibility = Visibility.Visible;
                    txtHintLogin.Visibility = Visibility.Visible;
                    PasswordBox.Visibility = Visibility.Visible;
                    txtHintPass.Visibility = Visibility.Visible;

                    ButtonChangePassword.Visibility = Visibility.Visible;
                    ButtonEnter.Visibility = Visibility.Visible;
                    ButtonReg.Visibility = Visibility.Visible;
                    return;

                case Visibility.Hidden:
                    // Показываем капчу и скрываем обычную форму
                    captcha.Visibility = Visibility.Visible;
                    captchaInput.Visibility = Visibility.Visible;
                    labelCaptcha.Visibility = Visibility.Visible;
                    submitCaptcha.Visibility = Visibility.Visible;

                    labelLogin.Visibility = Visibility.Hidden;
                    labelPass.Visibility = Visibility.Hidden;
                    TextBoxLogin.Visibility = Visibility.Hidden;
                    txtHintLogin.Visibility = Visibility.Hidden;
                    PasswordBox.Visibility = Visibility.Hidden;
                    txtHintPass.Visibility = Visibility.Hidden;

                    ButtonChangePassword.Visibility = Visibility.Hidden;
                    ButtonEnter.Visibility = Visibility.Hidden;
                    ButtonReg.Visibility = Visibility.Hidden;
                    return;
            }
        }

        public void CaptchaChange()
        {
            string allowchar = " ";
            allowchar = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z";
            allowchar += "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,y,z";
            allowchar += "1,2,3,4,5,6,7,8,9,0";
            char[] a = { ',' };
            string[] ar = allowchar.Split(a);
            string pwd = "";
            string temp = "";
            Random r = new Random();

            for (int i = 0; i < 6; i++)
            {
                temp = ar[(r.Next(0, ar.Length))];
                pwd += temp;
            }
            captcha.Text = pwd;
            captchaInput.Clear();
        }

        private void submitCaptcha_Click(object sender, RoutedEventArgs e)
        {
            if (captchaInput.Text != captcha.Text)
            {
                MessageBox.Show("Неверно введена капча", "Ошибка");
                CaptchaChange();
            }
            else
            {
                MessageBox.Show("Капча введена успешно, можете продолжить авторизацию", "Успех");
                CaptchaSwitch();
                failedAttempts = 0;
            }
        }

        private void textBox_PreviewExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Command == ApplicationCommands.Copy ||
                e.Command == ApplicationCommands.Cut ||
                e.Command == ApplicationCommands.Paste)
            {
                e.Handled = true;
            }
        }

        private void ButtonReg_Click(object sender, RoutedEventArgs e)
        {
            NavigationService?.Navigate(new RegPage());
        }

        private void ButtonChangePassword_Click(object sender, RoutedEventArgs e)
        {
            NavigationService?.Navigate(new ChangePassPage());
        }

        private void txtHintLogin_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            TextBoxLogin.Focus();
            txtHintLogin.Visibility = Visibility.Collapsed;
        }

        private void txtHintPass_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            PasswordBox.Focus();
            txtHintPass.Visibility = Visibility.Collapsed;
        }

        private void TextBoxLogin_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(TextBoxLogin.Text))
            {
                txtHintLogin.Visibility = Visibility.Visible;
            }
            else
            {
                txtHintLogin.Visibility = Visibility.Collapsed;
            }
        }

        private void PasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(PasswordBox.Password))
            {
                txtHintPass.Visibility = Visibility.Visible;
            }
            else
            {
                txtHintPass.Visibility = Visibility.Collapsed;
            }
        }
    }
}
