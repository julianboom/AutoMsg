using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Net;

namespace AutoMsg
{
    class FTPHelper
    {
        private static string FTPCONSTR = ConfigurationManager.AppSettings["FTPCONSTR"];//FTP的服务器地址，格式为ftp://192.168.1.234:8021/。ip地址和端口换成自己的，这些建议写在配置文件中，方便修改
        private static string FTPUSERNAME = ConfigurationManager.AppSettings["FTPUSERNAME"];//FTP服务器的用户名
        private static string FTPPASSWORD = ConfigurationManager.AppSettings["FTPPASSWORD"];//FTP服务器的密码
         /// <summary>
         /// 返回FTP服务器上的文件流
         /// </summary>
         /// <param name="ftpfilepath"></param>
         /// <returns>ftpStream</returns>
        public static Stream Download(string ftpfilepath)
        {
            Stream ftpStream = null;
            FtpWebResponse response = null;
            try
            {
                ftpfilepath = ftpfilepath.Replace("\\", "/");
                string url = FTPCONSTR + ftpfilepath;
                FtpWebRequest reqFtp = (FtpWebRequest)FtpWebRequest.Create(new Uri(url));
                reqFtp.UseBinary = true;
                reqFtp.Credentials = new NetworkCredential(FTPUSERNAME, FTPPASSWORD);
                response = (FtpWebResponse)reqFtp.GetResponse();
                ftpStream = response.GetResponseStream();
                
            }
            catch (Exception ee)
            {
                if (response != null)
                {
                    response.Close();
                }
                Console.WriteLine(ee);
            }
            return ftpStream;
        }
    }
}
