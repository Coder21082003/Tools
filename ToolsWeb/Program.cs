using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using OpenQA.Selenium.Interactions;
using System.Threading;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Support.UI;
using ClosedXML.Excel;
using System.IO;
using System.Diagnostics;
using DocumentFormat.OpenXml.Bibliography;
using System.Collections.ObjectModel;

namespace ToolsWeb
{

    public class Program
    {
        //Biến cần lưu
        private static string email, password, textToSpeech, ppScan, idScan, driverScan, temp = "";
        private static Random random = new Random();

        [STAThread]
        static void Main(string[] args)
        {
            //Tạo file excel trước khi vào vòng lặp 
            string apiKey = "ApiKey"; // Replace with your API Key or other suitable value
            string currentTime = DateTime.Now.ToString("ddMMyyyy");
            string filePath = $"../../../AccountData/{apiKey}_{currentTime}.xlsx";

            XLWorkbook workbook;
            if (File.Exists(filePath))
            {
                workbook = new XLWorkbook(filePath);
            }
            else
            {
                workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Data");
                worksheet.Cell(1, 1).Value = "Email";
                worksheet.Cell(1, 2).Value = "Password";
                worksheet.Cell(1, 3).Value = "TextToSpeech";
                worksheet.Cell(1, 4).Value = "PpScan";
                worksheet.Cell(1, 5).Value = "IdScan";
                worksheet.Cell(1, 6).Value = "DriverScan";
                workbook.SaveAs(filePath);
            }

            for (int i = 0; i < 1000; i++)
            {
                #region Khởi tạo driver

                ChromeOptions options = new ChromeOptions();
                options.AddArguments("--disable-notifications");

                // Khởi tạo ChromeDriver
                IWebDriver driver = new ChromeDriver(options);
                Actions actions = new Actions(driver);
                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;

                #endregion

                Thread.Sleep(1000);
                // Mở trang web thứ nhất - "https://mail1s.com/"
                driver.Navigate().GoToUrl("https://mail1s.com/");
                Thread.Sleep(800);
                

                //Chạy tools
                #region Đăng ký tk
                try
                {
                    IWebElement closeAd = null;
                    closeAd = driver.FindElement(By.XPath("//*[@id='hide_float_right']//a"));
                    if (closeAd != null)
                    {
                        closeAd.Click();
                    }

                    //Click để hiện qc
                    IWebElement create_rand_email = driver.FindElement(By.CssSelector("body > div > header > div > div.actions > div.p-3.md\\:p-0.w-full.md\\:w-2\\/4.order-2 > div.hidden.lg\\:block > div.app-action.mt-4.px-8 > div > div > form.flex.justify-center.mb-1 > a"));
                    create_rand_email.Click();

                    //Xóa quảng cáo
                    Thread.Sleep(1000);
                    IWebElement ads_positon_box = null;
                    ads_positon_box = driver.FindElement(By.CssSelector("html > ins"));
                    if (ads_positon_box != null)
                    {
                        jsExecutor.ExecuteScript("arguments[0].remove()", ads_positon_box);
                    }

                    //Tạo lại email
                    create_rand_email.Click();
                    Thread.Sleep(2000);

                    //Copy
                    IWebElement copy_btn = driver.FindElement(By.XPath("/html/body/div/header/div/div[3]/div[1]/div[2]/div[2]/div[2]/div[1]"));
                    copy_btn.Click();

                    // Mở tab thứ hai - "console.fpt.ai"
                    Thread.Sleep(1000);
                    jsExecutor.ExecuteScript("window.open('https://console.fpt.ai', '_blank');");

                    System.Threading.Thread.Sleep(2000);

                    var windowHandles = driver.WindowHandles;

                    driver.SwitchTo().Window(windowHandles[1]);

                    //Đăng kí tài khoản
                    IWebElement signUp_btn = driver.FindElement(By.CssSelector("#kc-registration > span > a"));
                    signUp_btn.Click();
                    Thread.Sleep(1000);

                    // Điền username, email
                    IWebElement userName = driver.FindElement(By.CssSelector("#username"));
                    IWebElement emailInput = driver.FindElement(By.CssSelector("#email"));

                    GetEmailValue();
                    userName.SendKeys(email);
                    emailInput.SendKeys(email);
                    Thread.Sleep(100);

                    //Điền password
                    IWebElement passWordInput = driver.FindElement(By.CssSelector("#password"));
                    IWebElement confirmPassWordInput = driver.FindElement(By.CssSelector("#password-confirm"));

                    password = GenerateRandomPassword();
                    passWordInput.SendKeys(password);
                    confirmPassWordInput.SendKeys(password);
                    Thread.Sleep(1000);

                    //Điền name
                    IWebElement firstName = driver.FindElement(By.CssSelector("#firstName"));
                    IWebElement lastName = driver.FindElement(By.CssSelector("#lastName"));

                    string fn = RandomNameGenerator.GenerateFirstName(random.Next(4, 8));
                    string ln = RandomNameGenerator.GenerateLastName(random.Next(4, 8));
                    firstName.SendKeys(fn);
                    lastName.SendKeys(ln);
                    Thread.Sleep(1000);

                    //Sign up
                    IWebElement signup_btn = driver.FindElement(By.CssSelector("#kc-signup-button"));
                    signup_btn.Click();
                    Thread.Sleep(1500);

                    //Sang trang mail1s để verify
                    driver.SwitchTo().Window(windowHandles[0]);
                    Thread.Sleep(1000);

                    bool checkEmailVerify = false;
                    IWebElement email_showup = null;
                    //Đợi email verify xuất hiện
                    Stopwatch stopwatch = new Stopwatch();
                    stopwatch.Start();
                    while (checkEmailVerify == false)
                    {
                        //Check Element exsist
                        try
                        {
                            if (stopwatch.ElapsedMilliseconds > 300000)
                            {
                                // Nếu thời gian vượt quá 30000ms (300 giây), đóng driver
                                stopwatch.Stop();
                                driver.Quit();
                                break;
                            }

                            email_showup = driver.FindElement(By.XPath("//main/div/div/div/div[2]/div[2]"));
                            if (email_showup != null)
                            {
                                checkEmailVerify = true;
                            }
                            break;
                        }
                        catch (NoSuchElementException)
                        {
                            Thread.Sleep(3000);
                            checkEmailVerify = false;
                        }

                    }

                    #endregion

                    #region Verify email

                    IWebElement divElement = driver.FindElement(By.XPath("//main/div/div/div/div[2]/div[2]"));
                    actions.MoveToElement(divElement).Perform();
                    divElement.Click();
                    Thread.Sleep(500);

                    IWebElement messageDiv = driver.FindElement(By.CssSelector("body > div > div.container.mx-auto.min-h-tm-half.order-2.bg-white.md\\:rounded-md.shadow-md.flex.space-x-2.justify-center.z-50.-mt-16.-mb-16.ct16 > main > div > div > div.message"));
                    IWebElement iframeEmail = messageDiv.FindElement(By.XPath(".//div[3]//iframe"));
                    driver.SwitchTo().Frame(iframeEmail);

                    IWebElement verify_btn = driver.FindElement(By.XPath("//td[contains(@class, 'center_content')]/div[contains(@class, 'editable-text')]/span[contains(@class, 'text_container')]/a"));
                    actions.MoveToElement(verify_btn).Perform();
                    verify_btn.Click();

                    //Đóng cửa sổ kia 
                    driver.SwitchTo().Window(windowHandles[1]);
                    driver.Close();

                    #endregion


                    #region Lấy các api key 
                    windowHandles = driver.WindowHandles;
                    driver.SwitchTo().Window(windowHandles[1]);

                    //Đợi trang web load xong
                    Thread.Sleep(2500);
                    IWebElement overlay = driver.FindElement(By.TagName("body"));
                    overlay.Click();


                    #region Lấy text to speech api key

                    IWebElement text_to_speech_btn = driver.FindElement(By.XPath("//div[@class='box-wrapper']//div[1]/div[1]/div/div/div/div/div[1]/div"));
                    text_to_speech_btn.Click();

                    IWebElement confirm_text_to_speech_btn = driver.FindElement(By.CssSelector("#app > div.v-dialog__content.v-dialog__content--active > div > div > div.v-card__actions > button.v-btn.theme--light.primary"));
                    confirm_text_to_speech_btn.Click();

                    Thread.Sleep(1000);
                    IWebElement project_name = driver.FindElement(By.XPath("//*[@id=\"app\"]/div[12]/div/div/div[2]/form/div[2]/div/div[1]/div/input"));
                    string pjName = RandomNameGenerator.GenerateFirstName(3);
                    Thread.Sleep(200);
                    project_name.SendKeys(pjName);

                    IWebElement create_text_to_speech_btn = driver.FindElement(By.CssSelector("#app > div.v-dialog__content.v-dialog__content--active > div > div > div.v-card__actions > button:nth-child(2)"));
                    create_text_to_speech_btn.Click();
                    Thread.Sleep(1000);

                    IWebElement api_key_label = driver.FindElement(By.XPath("//*[@id=\"app\"]/div[11]/div/div/div[2]/form/div/div/div[1]/div/input"));
                    api_key_label.SendKeys(pjName);

                    IWebElement create_txsp_api_key = driver.FindElement(By.CssSelector("#app > div.v-dialog__content.v-dialog__content--active > div > div > div.v-card__actions > button.v-btn.v-btn--flat.theme--light.warning--text.text--darken-1"));
                    create_txsp_api_key.Click();

                    Thread.Sleep(1200);
                    IWebElement code_ttsp = driver.FindElement(By.XPath("//*[@id=\"fptai-tts\"]/div[2]/div[2]/code"));
                    temp = code_ttsp.Text.Trim();
                    textToSpeech = ExtractApiKey(temp);
                    temp = "";
                    Thread.Sleep(500);

                    #endregion

                    #region Lấy passport api key

                    //Back lại apis
                    driver.Navigate().Back();
                    Thread.Sleep(1000);
                    overlay.Click();

                    //create passport api
                    IWebElement create_passport_apis_btn = driver.FindElement(By.XPath("//div[@class='box-wrapper']//div[3]/div[1]/div/div/div/div/div[1]/div"));
                    create_passport_apis_btn.Click();
                    Thread.Sleep(200);

                    IWebElement confirm_passport_api = driver.FindElement(By.CssSelector("#app > div.v-dialog__content.v-dialog__content--active > div > div > div.v-card__actions > button.v-btn.theme--light.primary"));
                    confirm_passport_api.Click();
                    Thread.Sleep(1000);

                    //Take passport api key
                    IWebElement code_passport = driver.FindElement(By.XPath("//*[@id=\"app\"]/div[15]/div[1]/main/div/div/div/div[2]/div[2]/div/div[2]/div/div[2]/code"));
                    temp = code_passport.Text.Trim();
                    ppScan = ExtractApiKey(temp);
                    temp = "";
                    Thread.Sleep(500);

                    #endregion


                    #region Lấy idRecognition api key

                    //Back lại apis
                    driver.Navigate().Back();
                    Thread.Sleep(1000);
                    overlay.Click();

                    //create idRecognition api
                    IWebElement create_idRecognition_apis_btn = driver.FindElement(By.XPath("//div[@class='box-wrapper']//div[4]/div[1]/div/div/div/div/div[1]/div"));
                    create_idRecognition_apis_btn.Click();
                    Thread.Sleep(200);

                    IWebElement confirm_idRecognition_api = driver.FindElement(By.CssSelector("#app > div.v-dialog__content.v-dialog__content--active > div > div > div.v-card__actions > button.v-btn.theme--light.primary"));
                    confirm_idRecognition_api.Click();
                    Thread.Sleep(1000);

                    //Take idRecognition api key
                    IWebElement codeidRecognition = driver.FindElement(By.XPath("//*[@id=\"app\"]/div[15]/div[1]/main/div/div/div/div[2]/div[2]/div/div[2]/div/div[2]/code"));
                    temp = codeidRecognition.Text.Trim();
                    idScan = ExtractApiKey(temp);
                    temp = "";
                    Thread.Sleep(500);

                    #endregion


                    #region Lấy ID driver license api key

                    //Back lại apis
                    driver.Navigate().Back();
                    Thread.Sleep(1000);
                    overlay.Click();

                    //Create driver license api key
                    IWebElement create_driverLicense_apis_btn = driver.FindElement(By.XPath("//div[@class='box-wrapper']//div[6]/div[1]/div/div/div/div/div[1]/div"));
                    create_driverLicense_apis_btn.Click();
                    Thread.Sleep(200);

                    IWebElement confirm_driverLicense_api = driver.FindElement(By.CssSelector("#app > div.v-dialog__content.v-dialog__content--active > div > div > div.v-card__actions > button.v-btn.theme--light.primary"));
                    confirm_driverLicense_api.Click();
                    Thread.Sleep(1000);

                    //Lấy key
                    IWebElement codeDriverLicense = driver.FindElement(By.XPath("//*[@id=\"app\"]/div[15]/div[1]/main/div/div/div/div[2]/div[2]/div/div[2]/div/div[2]/code"));
                    temp = codeDriverLicense.Text.Trim();
                    driverScan = ExtractApiKey(temp);
                    temp = "";
                    Thread.Sleep(500);

                    //Xuất data ra file excel 
                    AddDataToExcel(workbook, email, password, textToSpeech, ppScan, idScan, driverScan);


                    // Đóng driver
                    driver.Close();
                    driver.Quit();
                    #endregion
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Đã xảy ra lỗi: " + ex.Message);
                    driver.Quit();
                }
                #endregion
            }
        }

        #region Các method

        //Lưu file excel
        static void AddDataToExcel(XLWorkbook workbook, string email, string password, string textToSpeech, string ppScan, string idScan, string driverScan)
        {
            // Get the worksheet with the name "Data"
            var worksheet = workbook.Worksheet("Data");

            // Get the row count to determine where to add the new data
            int rowCount = worksheet.RowsUsed().Count() + 1;

            // Add data to the next row
            worksheet.Cell(rowCount, 1).Value = email;
            worksheet.Cell(rowCount, 2).Value = password;
            worksheet.Cell(rowCount, 3).Value = textToSpeech;
            worksheet.Cell(rowCount, 4).Value = ppScan;
            worksheet.Cell(rowCount, 5).Value = idScan;
            worksheet.Cell(rowCount, 6).Value = driverScan;

            workbook.Save();
            workbook.Dispose();
        }

        //Regex api key
        public static string ExtractApiKey(string input)
        {
            string pattern = "(?<=api-key: )([^\"]+)";

            Match match = Regex.Match(input, pattern);

            if (match.Success)
            {
                string apiKey = match.Groups[1].Value;
                return apiKey;
            }
            else
            {
                return null; // Hoặc giá trị mặc định khác tùy ý
            }
        }

        //get clipboarvalue
        public static string GetEmailValue()
        {
            email = Clipboard.GetText();
            return email;
        }
        //Random password
        public static string GenerateRandomPassword()
        {
            string upperChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string lowerChars = "abcdefghijklmnopqrstuvwxyz";
            string digits = "0123456789";
            string specialChars = "!@#$%^&*()_+-=[]{}|;:,.<>?";

            // Tạo một StringBuilder để xây dựng password
            StringBuilder passwordBuilder = new StringBuilder();

            // Sử dụng lớp Random để tạo số ngẫu nhiên
            Random random = new Random();

            // Chọn ít nhất một ký tự hoa, một ký tự thường, một chữ số và một ký tự đặc biệt
            passwordBuilder.Append(upperChars[random.Next(upperChars.Length)]);
            passwordBuilder.Append(lowerChars[random.Next(lowerChars.Length)]);
            passwordBuilder.Append(digits[random.Next(digits.Length)]);
            passwordBuilder.Append(specialChars[random.Next(specialChars.Length)]);

            // Tiếp tục thêm các ký tự ngẫu nhiên cho đến khi password đạt đủ 12 ký tự
            int remainingLength = 12 - passwordBuilder.Length;
            for (int i = 0; i < remainingLength; i++)
            {
                string allChars = upperChars + lowerChars + digits + specialChars;
                passwordBuilder.Append(allChars[random.Next(allChars.Length)]);
            }

            // Trộn ngẫu nhiên các ký tự trong password
            string password = ShufflePassword(passwordBuilder.ToString());

            return password;
        }
        public static string ShufflePassword(string password)
        {
            // Chuyển password thành mảng ký tự để trộn ngẫu nhiên
            char[] passwordArray = password.ToCharArray();
            Random random = new Random();

            // Trộn ngẫu nhiên mảng ký tự
            for (int i = 0; i < passwordArray.Length - 1; i++)
            {
                int j = random.Next(i, passwordArray.Length);
                char temp = passwordArray[i];
                passwordArray[i] = passwordArray[j];
                passwordArray[j] = temp;
            }

            // Chuyển mảng ký tự thành chuỗi password mới
            string shuffledPassword = new string(passwordArray);

            return shuffledPassword;
        }


        //Random name
        public class RandomNameGenerator
        {
            private static Random random = new Random();

            // Hàm tạo ngẫu nhiên tên
            public static string GenerateFirstName(int length)
            {
                string characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
                return GenerateRandomString(characters, length);
            }

            // Hàm tạo ngẫu nhiên họ
            public static string GenerateLastName(int length)
            {
                string characters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
                return GenerateRandomString(characters, length);
            }

            // Hàm tạo chuỗi ngẫu nhiên từ danh sách kí tự
            private static string GenerateRandomString(string characters, int length)
            {
                char[] result = new char[length];
                for (int i = 0; i < length; i++)
                {
                    result[i] = characters[random.Next(characters.Length)];
                }
                return new string(result);
            }
        }
        #endregion
    }
}
