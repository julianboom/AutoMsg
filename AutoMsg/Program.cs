using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Cells;
using System.Configuration;
using System.Timers;

namespace AutoMsg
{
    class Program
    {        
        static void Main(string[] args)
        {
            MsgService msgService = new MsgService();
            msgService.IntervalClock();//开启定时任务
            Console.ReadKey();
        }

        



    }

    /// <summary>
    /// 短信发送服务类
    /// </summary>
    public class MsgService
    {
        System.Threading.Timer timer = null;

        public void SendMsg(object o)
        {
            Console.WriteLine($"{DateTime.Now},短信自动发送服务正在运行中...");
            ConfigurationManager.RefreshSection("appSettings");
            string intervalTime = ConfigurationManager.AppSettings["INTERVALTIME"]??"1";
            DateTime start = DateTime.Parse(ConfigurationManager.AppSettings["STARTTIME"]??"00:00");//获取服务时间区间
            DateTime end = DateTime.Parse(ConfigurationManager.AppSettings["ENDTIME"]??"00:00");//获取服务时间区间
            DateTime now = DateTime.Parse(DateTime.Now.ToString("HH:mm"));//获取服务时间区间
            string fileName = ConfigurationManager.AppSettings["FILENAME"]??"messages.xls";
            //log("服务时间段:" + start + "--" + end + "  间隔时间:" + intervalTime + "Minutes");

            timer.Change((Convert.ToInt32(intervalTime) * 60 * 1000), 0);//自动根据配置文件更新定时器时间
            if (start < now && now < end)
            {
            Stream stream= FTPHelper.Download(fileName);
            if(stream == null)
                {
                    log("无法获取FTP服务器文件，请检查配置信息是否正确以及文件是否存在");
                    this.SendMsg(null);//当报错时一直循环请求文件，直到成功获取到文件流
                    return;
                }
            Stream streamout = new MemoryStream();
            stream.CopyTo(streamout);
            Workbook workbook = new Workbook();
            workbook.Open(streamout, FileFormatType.Excel97To2003);
            Cells cells = workbook.Worksheets[0].Cells;
            String[,] array = new String[cells.MaxDataRow + 1, cells.MaxDataColumn + 1];
            for (int i = 0; i < cells.MaxDataRow + 1; i++)
            {
                for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                {

                    string s = cells[i, j].StringValue.Trim();
                    array[i, j] = s;
                }
            }
                string tempMobilPhone = "";
                string tempContent = "";
            for (int i = 1; i < cells.MaxDataRow + 1; i++)
            {
                for (int j = 1; j < cells.MaxDataColumn + 1; j++)
                {
                    if (j == 1)
                    {
                            tempContent = array[i, j];

                    }
                    else
                    {
                            tempMobilPhone = array[i, j];

                    }
                }
                    string result = CMCC(tempMobilPhone, tempContent);
                    if (result != "false")
                    {
                        log(tempMobilPhone+": "+tempContent + " 短信已发出;" + result);
                    }
                }
                Console.WriteLine($"{DateTime.Now},短信发送成功！");
            }
            else
            {
                log("服务时间段:" + start + "--" + end + "  间隔时间:" + intervalTime + "Minutes\r\n当前不在服务时间段");
                Console.WriteLine($"{DateTime.Now},当前不在服务时间段");
            }



        }
        /// <summary>
        /// CMCC短信服务
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static string CMCC(string mobilePhone, string msgContent)
        {
            ConfigurationManager.RefreshSection("appSettings");
            string ws_addr = ConfigurationManager.AppSettings["WS_ADDR"] ?? "";
            string ws_appid = ConfigurationManager.AppSettings["WS_APPID"] ?? "";

            try
            {
                cmcc_mas_wbs wbs = new cmcc_mas_wbs();
                wbs.Url = ws_addr;
                sendSmsRequest smsreq = new sendSmsRequest();
                smsreq.ApplicationID = ws_appid;
                //smsreq.DeliveryResultRequest = true;
                smsreq.DestinationAddresses = new string[] { "tel:" + mobilePhone };//手机号码
                smsreq.Message = msgContent;//短信内容
                smsreq.MessageFormat = MessageFormat.GB2312;
                smsreq.SendMethod = SendMethodType.Long;
                sendSmsResponse rsp = wbs.sendSms(smsreq);
                string sendResultID = rsp.RequestIdentifier;//返回的唯一标识符
                //GetSmsDeliveryStatusRequest GetSmsDeliveryStatusRequest = new GetSmsDeliveryStatusRequest();
                //GetSmsDeliveryStatusRequest.ApplicationID = ws_appid;
                //GetSmsDeliveryStatusRequest.ApplicationID = sendResultID;

                return (sendResultID /*+"&"+ sendResultID*/);
            }
            catch (Exception ex)
            {
                log("短信服务出错");
                return ("false");
            }
        }

        /// <summary>
        /// 定时任务入口
        /// </summary>
        public void IntervalClock()
        {
            timer = new System.Threading.Timer(SendMsg, null, 1000, 0);
        }


        /// <summary>
        /// 日志工具
        /// </summary>
        public static void log(string content)
        {
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "MyServiceLog.txt";
            FileStream stream = null;
            StreamWriter writer = null;

            string template = "\r\n{0}\r\n{1}";
            try
            {
                stream = new FileStream(filePath, FileMode.Append);
                writer = new StreamWriter(stream);
                writer.WriteLine(string.Format(template, DateTime.Now, content));
                writer.Close();
                stream.Close();
            }
            catch(Exception e)
            {
                Console.WriteLine("日志写出错误，请联系管理员及时处理");
            }
            finally
            {
                if(writer != null)
                {
                    writer.Close();
                }
                if(stream != null)
                {
                    stream.Close();
                }
            }

        }
    }
}
