using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace BingBackground {

    class BingBackground {

        private static void Main(string[] args) {
            //Console.WriteLine("APP version: " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString());
            string imgFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds");
            Directory.CreateDirectory(imgFolder);
            string urlBase = GetBackgroundUrlBase();
            if(urlBase != "")
            {
                Image background = DownloadBackground(urlBase);
                if (background != null)
                {
                    string imgPath = SaveBackground(background, imgFolder);
                    DelImgExpired(7, imgFolder);
                    int RetryRound = 0;
                    while (RetryRound < 5 && SetBackground(imgPath, GetPosition()) == false)
                    {
                        RetryRound++;
                        if (RetryRound == 5)
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine("Press enter to exit...");
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Console.ReadLine();
                        }
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("No img found...");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Press enter to exit...");
                    Console.ReadLine();
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("No img found...");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Press enter to exit...");
                Console.ReadLine();
            }
            
        }

        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }

        private static dynamic DownloadJson() {
            using (WebClient webClient = new WebClient()) {
                Console.WriteLine("Downloading JSON...");
                try
                {
                    /*
                    请求某些接口一直返回基础连接已关闭：发送时发生错误
                    远程主机强迫关闭了一个现有的连接
                    */
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                    ServicePointManager.Expect100Continue = false;
                    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);

                    //Resolve the encoding issue for Chinese
                    webClient.Encoding = System.Text.Encoding.UTF8;
                    string jsonString = webClient.DownloadString("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US");
                    return JsonConvert.DeserializeObject<dynamic>(jsonString);
                }
                catch (Exception e)  
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Cannot get json from the url because: \n" + e.Message);
                    return null;
                }
            }
        }

        private static string GetBackgroundUrlBase() {
            dynamic jsonObject = DownloadJson();
            if(jsonObject != null)
            {
                return "https://www.bing.com" + jsonObject.images[0].url;
            }
            else
            {
                return "";
            }
        }

        private static string GetBackgroundTitle() {
            dynamic jsonObject = DownloadJson();
            string copyrightText = jsonObject.images[0].copyright;
            return copyrightText.Substring(0, copyrightText.IndexOf(" ("));
        }

        private static bool WebsiteExists(string url) {
            try {
                WebRequest request = WebRequest.Create(url);
                request.Method = "HEAD";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                return response.StatusCode == HttpStatusCode.OK;
            } catch {
                return false;
            }
        }

        private static string GetResolutionExtension(string url) {
            Rectangle resolution = Screen.PrimaryScreen.Bounds;
            string widthByHeight = resolution.Width + "x" + resolution.Height;
            string potentialExtension = "_" + widthByHeight + ".jpg";
            if (WebsiteExists(url + potentialExtension)) {
                Console.WriteLine("Background for " + widthByHeight + " found.");
                return potentialExtension;
            } else {
                Console.WriteLine("No background for " + widthByHeight + " was found.");
                Console.WriteLine("Using 1920x1080 instead.");
                return "_1920x1080.jpg";
            }
        }

        private static void SetProxy() {
            string proxyUrl = Properties.Settings.Default.Proxy;
            if (proxyUrl.Length > 0) {
                WebProxy webProxy = new WebProxy(proxyUrl, true);
                webProxy.Credentials = CredentialCache.DefaultCredentials;
                WebRequest.DefaultWebProxy = webProxy;     
            }
        }

        private static Image DownloadBackground(string url) {
            Console.WriteLine("Downloading background...");
            //SetProxy();
            if (WebsiteExists(url))
            {
                WebRequest request = WebRequest.Create(url);
                WebResponse reponse = request.GetResponse();
                Stream stream = reponse.GetResponseStream();
                return Image.FromStream(stream);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Cannot download picutre from: " + url);
                return null;
            }
        }

        private static string SaveBackground(Image Background, string FolderName) {
            Console.WriteLine("Saving background...");
            string imgPath = Path.Combine(FolderName, DateTime.Now.ToString("M-d-yyyy") + ".bmp");
            Background.Save(imgPath, System.Drawing.Imaging.ImageFormat.Bmp);
            Console.WriteLine("Image saved as: " + imgPath);
            return imgPath;
        }
        
        private enum PicturePosition {
            Tile,
            Center,
            Stretch,
            Fit,
            Fill
        }

        private static PicturePosition GetPosition() {
            PicturePosition position = PicturePosition.Fit;
            switch (Properties.Settings.Default.Position) {
                case "Tile":
                    position = PicturePosition.Tile;
                    break;
                case "Center":
                    position = PicturePosition.Center;
                    break;
                case "Stretch":
                    position = PicturePosition.Stretch;
                    break;
                case "Fit":
                    position = PicturePosition.Fit;
                    break;
                case "Fill":
                    position = PicturePosition.Fill;
                    break;
            }
            return position;
        }

        internal sealed class NativeMethods {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            internal static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }

        private static Boolean SetBackground(string ImgFullName, PicturePosition style) {
            Console.WriteLine("Setting background...");
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(Path.Combine("Control Panel", "Desktop"), true)) {
                switch (style) {
                    case PicturePosition.Tile:
                        key.SetValue("WallpaperStyle", "0");
                        key.SetValue("TileWallpaper", "1");
                        break;
                    case PicturePosition.Center:
                        key.SetValue("WallpaperStyle", "0");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Stretch:
                        key.SetValue("WallpaperStyle", "2");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Fit:
                        key.SetValue("WallpaperStyle", "6");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Fill:
                        key.SetValue("WallpaperStyle", "10");
                        key.SetValue("TileWallpaper", "0");
                        break;
                }
            }
            const int SetDesktopBackground = 20;
            const int UpdateIniFile = 1;
            const int SendWindowsIniChange = 2;
            int NumResult = NativeMethods.SystemParametersInfo(SetDesktopBackground, 0, ImgFullName, UpdateIniFile | SendWindowsIniChange);
            if (NumResult == 1)
            {
                Console.WriteLine("Background updated succesfully!");
                return true;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Background update failed...");
                Console.ForegroundColor = ConsoleColor.Gray;
                return false;
            }
        }

        private static void DelImgExpired(int days, string FolderName)
        {
            foreach (string FileName in Directory.EnumerateFiles(FolderName))
            {
                if ((DateTime.Today.Date - Directory.GetCreationTime(FileName).Date).Days >= days)
                {
                    File.Delete(FileName);
                    Console.WriteLine("Deleted image: " + FileName);
                }
            }
        }

    }

}
