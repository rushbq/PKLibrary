using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using ClosedXML.Excel;
using System.Net.Mail;

namespace PKLib_Method.Methods
{
    public static class CustomExtension
    {
        #region -- 一般功能 --
        /// <summary>
        /// 簡化string.format
        /// </summary>
        /// <param name="format"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static string FormatThis(this string format, params object[] args)
        {
            return string.Format(format, args);
        }

        /// <summary>
        /// 取得Right字串
        /// </summary>
        /// <param name="inputValue">輸入字串</param>
        /// <param name="length">取得長度</param>
        /// <returns>string</returns>
        /// <example>
        /// string str = "12345";
        /// str = str.Right(3);  //345
        /// </example>
        public static string Right(this string inputValue, int length)
        {
            length = Math.Max(length, 0);

            if (inputValue.Length > length)
            {
                return inputValue.Substring(inputValue.Length - length, length);
            }
            else
            {
                return inputValue;
            }
        }

        /// <summary>
        /// 取得Left字串
        /// </summary>
        /// <param name="inputValue">輸入字串</param>
        /// <param name="length">取得長度</param>
        /// <returns>string</returns>
        /// <example>
        /// string str = "12345";
        /// str = str.Left(3);  //123
        /// </example>
        public static string Left(this string inputValue, int length)
        {
            length = Math.Max(length, 0);

            if (inputValue.Length > length)
            {
                return inputValue.Substring(0, length);
            }
            else
            {
                return inputValue;
            }
        }


        /// <summary>
        /// 金額格式轉換 (含三位點)
        /// </summary>
        /// <param name="inputValue">傳入的值</param>
        /// <returns>string</returns>
        /// <example>550.ToMoneyString()</example>
        public static string ToMoneyString(this string inputValue)
        {
            try
            {
                //去除三位點
                inputValue = inputValue.Replace(",", "");

                //判斷是否為數值
                if (inputValue.IsNumeric() == false)
                    return inputValue;

                //轉型為Double
                double dbl_Value = Convert.ToDouble(inputValue);
                //金額 >= 1000
                if (dbl_Value >= 1000)
                {
                    if (dbl_Value > Math.Floor(dbl_Value))
                        return String.Format("{0:#,000.00}", dbl_Value);
                    else
                        return String.Format("{0:#,000}", dbl_Value);
                }
                //金額 > 0 And 金額 < 1000
                if (dbl_Value > 0 & dbl_Value < 1000)
                {
                    if (dbl_Value > Math.Floor(dbl_Value))
                        return String.Format("{0:0.00}", dbl_Value);
                    else
                        return Convert.ToString(dbl_Value);
                }
                //金額 = 0
                if (dbl_Value == 0)
                    return Convert.ToString(dbl_Value);

                //金額 < 0 And 金額 > -1000
                if (dbl_Value < 0 & dbl_Value > -1000)
                {
                    if (Math.Abs(dbl_Value) > Math.Floor(Math.Abs(dbl_Value)))
                        return String.Format("-{0:0.00}", Math.Abs(dbl_Value));
                    else
                        return String.Format("-{0}", Math.Abs(dbl_Value));
                }

                //金額 < -1000
                if (dbl_Value < -1000)
                {
                    if (Math.Abs(dbl_Value) > Math.Floor(Math.Abs(dbl_Value)))
                        return String.Format("-{0:#,000.00}", Math.Abs(dbl_Value));
                    else
                        return String.Format("-{0:#,000}", Math.Abs(dbl_Value));
                }

                return inputValue;
            }
            catch (Exception)
            {
                return "0";

            }
        }


        /// <summary>
        /// 數字小數點格式轉換(四捨五入)
        /// </summary>
        /// <param name="inputValue">傳入的值</param>
        /// <param name="idxNumber">取到第幾位</param>
        /// <returns>string</returns>
        public static string ToDecimalString(this string inputValue, int idxNumber)
        {
            try
            {
                if (string.IsNullOrEmpty(inputValue))
                    return "";
                if (inputValue.IsNumeric() == false)
                    return "";
                if (idxNumber < 0)
                    return "";

                return Math.Round(Convert.ToDouble(inputValue), idxNumber, MidpointRounding.AwayFromZero).ToString();
            }
            catch (Exception)
            {
                return "0";
            }
        }


        /// <summary>
        /// 取得各參數串的值
        /// </summary>
        /// <param name="str">String to process</param>
        /// <param name="OuterSeparator">Separator for each "NameValue"</param>
        /// <param name="NameValueSeparator">Separator for Name/Value splitting</param>
        /// <returns></returns>
        /// <example>
        /// string para = "param1=value1;param2=value2";
        /// NameValueCollection _data = para.ToNameValueCollection(';', '=');
        /// foreach (var item in _data.AllKeys)
        /// {
        ///     string _name = item;
        ///     string _val = _data[item];
        /// }
        /// </example>
        public static NameValueCollection ToNameValueCollection(this String inputValue, Char OuterSeparator, Char NameValueSeparator)
        {
            NameValueCollection nvText = null;
            inputValue = inputValue.TrimEnd(OuterSeparator);
            if (!String.IsNullOrEmpty(inputValue))
            {
                String[] arrStrings = inputValue.TrimEnd(OuterSeparator).Split(OuterSeparator);

                foreach (String s in arrStrings)
                {
                    Int32 posSep = s.IndexOf(NameValueSeparator);
                    String name = s.Substring(0, posSep);
                    String value = s.Substring(posSep + 1);
                    if (nvText == null)
                        nvText = new NameValueCollection();
                    nvText.Add(name, value);
                }
            }
            return nvText;
        }

        /// <summary>
        /// 檢查格式 - 是否為日期
        /// </summary>
        /// <param name="inputValue">日期</param>
        /// <returns>bool</returns>
        /// <example>
        /// string someDate = "2010/1/5";
        /// bool isDate = nonDate.IsDate();
        /// </example>
        public static bool IsDate(this string inputValue)
        {
            if (!string.IsNullOrEmpty(inputValue))
            {
                DateTime dt;
                return (DateTime.TryParse(inputValue, out dt));
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 檢查格式 - 是否為網址
        /// </summary>
        /// <param name="inputValue">網址字串</param>
        /// <returns>bool</returns>
        public static bool IsUrl(this string inputValue)
        {
            return Regex.IsMatch(inputValue, @"^(ht|f)tp(s?)\:\/\/[0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*(:(0-9)*)*(\/?)([a-zA-Z0-9\-\.\?\,\'\/\\\+&amp;%\$#_]*)?$");
        }

        /// <summary>
        /// 檢查格式 - 是否為座標
        /// </summary>
        /// <param name="Lat">座標-Lat字串</param>
        /// <param name="Lng">座標-Lng字串</param>
        /// <returns>Boolean</returns>
        public static bool IsLatLng(string Lat, string Lng)
        {
            if (IsNumeric(Lat) & IsNumeric(Lng))
            {
                if (Math.Abs(Convert.ToDouble(Lat)) >= 0 & Math.Abs(Convert.ToDouble(Lat)) < 180 & Math.Abs(Convert.ToDouble(Lng)) >= 0 & Math.Abs(Convert.ToDouble(Lng)) < 180)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 檢查格式 - 是否為數字
        /// </summary>
        /// <param name="Expression">輸入值</param>
        /// <returns>bool</returns>
        /// <see cref="http://support.microsoft.com/kb/329488/zh-tw"/>
        public static bool IsNumeric(this object Expression)
        {
            // Variable to collect the Return value of the TryParse method.
            bool isNum;
            // Define variable to collect out parameter of the TryParse method. If the conversion fails, the out parameter is zero.
            double retNum;
            // The TryParse method converts a string in a specified style and culture-specific format to its double-precision floating point number equivalent.
            // The TryParse method does not generate an exception if the conversion fails. If the conversion passes, True is returned. If it does not, False is returned.
            isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);

            return isNum;
        }

        /// <summary>
        /// 檢查格式 - EMail
        /// </summary>
        /// <param name="inputValue">Email</param>
        /// <returns>bool</returns>
        public static bool IsEmail(this string inputValue)
        {
            // Return true if strIn is in valid e-mail format.
            return Regex.IsMatch(inputValue,
                   @"^(?("")("".+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +
                   @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,6}))$");
        }

        /// <summary>
        /// 日期格式化
        /// </summary>
        /// <param name="inputValue">日期字串</param>
        /// <param name="formatValue">要輸出的格式</param>
        /// <returns>string</returns>
        public static string ToDateString(this string inputValue, string formatValue)
        {
            if (string.IsNullOrEmpty(inputValue))
            {
                return "";
            }
            else
            {
                return String.Format("{0:" + formatValue + "}", Convert.ToDateTime(inputValue));
            }

        }

        /// <summary>
        /// 日期格式化 - ERP
        /// 將ERP字串格式的日期,輸出成正常日期格式
        /// </summary>
        /// <param name="inputValue">日期字串</param>
        /// <param name="stringIcon">日期間隔符號</param>
        /// <returns>string</returns>
        /// <example>
        /// 原始日期:20101215
        /// "20201231".ToDateString_ERP("/") = 2020/12/31
        /// </example>
        public static string ToDateString_ERP(this string inputValue, string stringIcon)
        {
            if (string.IsNullOrEmpty(inputValue))
            {
                return "";
            }
            else
            {
                return String.Format("{1}{0}{2}{0}{3}"
                    , stringIcon
                    , inputValue.Substring(0, 4)
                    , inputValue.Substring(4, 2)
                    , inputValue.Substring(6, 2));
            }

        }


        /// <summary>
        /// 產生隨機英數字
        /// </summary>
        /// <param name="VcodeNum">顯示幾碼</param>
        /// <returns>string</returns>
        public static string GetRndNum(int VcodeNum)
        {
            string Vchar = "a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,u,v,w,x,y,z,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,0,1,2,3,4,5,6,7,8,9";
            string[] VcArray = Vchar.Split(',');
            string VNum = ""; //由于字符串很短，就不用StringBuilder了
            int temp = -1; //记录上次随机数值，尽量避免生产几个一样的随机数
            //采用一个简单的算法以保证生成随机数的不同
            Random rand = new Random();
            for (int i = 1; i < VcodeNum + 1; i++)
            {
                if (temp != -1)
                {
                    rand = new Random(i * temp * unchecked((int)DateTime.Now.Ticks));
                }
                int t = rand.Next(VcArray.Length);
                if (temp != -1 && temp == t)
                {
                    return GetRndNum(VcodeNum);
                }
                temp = t;
                VNum += VcArray[t];
            }
            return VNum;
        }

        /// <summary>
        /// 取得IP
        /// </summary>
        /// <returns></returns>
        public static string GetIP()
        {
            string ip;
            string trueIP = string.Empty;
            HttpRequest req = HttpContext.Current.Request;

            //先取得是否有經過代理伺服器
            ip = req.ServerVariables["HTTP_X_FORWARDED_FOR"];

            if (!string.IsNullOrEmpty(ip))
            {
                //將取得的 IP 字串存入陣列
                string[] ipRange = ip.Split(',');

                //比對陣列中的每個 IP
                for (int i = 0; i < ipRange.Length; i++)
                {
                    //剔除內部 IP 及不合法的 IP 後，取出第一個合法 IP
                    if (ipRange[i].Trim().Substring(0, 3) != "10." &&
                        ipRange[i].Trim().Substring(0, 7) != "192.168" &&
                        ipRange[i].Trim().Substring(0, 7) != "172.16." &&
                        CheckIP(ipRange[i].Trim()))
                    {
                        trueIP = ipRange[i].Trim();
                        break;
                    }
                }

            }
            else
            {
                //沒經過代理伺服器，直接使用 ServerVariables["REMOTE_ADDR"]
                //並經過 CheckIP( ) 的驗證
                trueIP = CheckIP(req.ServerVariables["REMOTE_ADDR"]) ?
                    req.ServerVariables["REMOTE_ADDR"] : "";
            }

            return trueIP;
        }

        /// <summary>
        /// 檢查 IP 是否合法
        /// </summary>
        /// <param name="strPattern">需檢測的 IP</param>
        /// <returns>true:合法 false:不合法</returns>
        private static bool CheckIP(string strPattern)
        {
            // 繼承自：System.Text.RegularExpressions
            // regular: ^\d{1,3}[\.]\d{1,3}[\.]\d{1,3}[\.]\d{1,3}$
            Regex regex = new Regex("^\\d{1,3}[\\.]\\d{1,3}[\\.]\\d{1,3}[\\.]\\d{1,3}$");
            Match m = regex.Match(strPattern);

            return m.Success;
        }

        /// <summary>
        /// 建立Url
        /// </summary>
        /// <param name="Uri">網址</param>
        /// <param name="ParamName">參數名稱(Array)(String)</param>
        /// <param name="ParamVal">參數值(Array)(String)</param>
        /// <returns>string</returns>
        public static string CreateUrl(string Uri, Array ParamName, Array ParamVal)
        {
            //判斷網址是否為空
            if (string.IsNullOrEmpty(Uri))
            {
                return "";
            }

            //產生完整網址
            StringBuilder url = new StringBuilder();
            url.Append(Uri);

            //判斷陣列索引數是否相同
            if (ParamName.Length == ParamVal.Length)
            {
                for (int row = 0; row < ParamName.Length; row++)
                {
                    url.Append(string.Format("{0}{1}={2}"
                        , (row == 0) ? "?" : "&"
                        , ParamName.GetValue(row).ToString()
                        , HttpUtility.UrlEncode(ParamVal.GetValue(row).ToString())
                        ));
                }
            }

            return url.ToString();
        }

        public static string CreateUrl(string Uri, Dictionary<string, string> _params)
        {
            //判斷網址是否為空
            if (string.IsNullOrEmpty(Uri))
            {
                return "";
            }

            //產生完整網址
            StringBuilder url = new StringBuilder();
            url.Append(Uri);

            int row = 0;
            foreach (var item in _params)
            {
                string pName = item.Key;
                string pValue = item.Value;

                url.Append(string.Format("{0}{1}={2}"
                    , (row == 0) ? "?" : "&"
                    , pName
                    , HttpUtility.UrlEncode(pValue)
                    ));

                row++;
            }

            return url.ToString();
        }

        /// <summary>
        /// 判斷字串內是否包含某字詞
        /// </summary>
        /// <param name="inputValue">輸入字串</param>
        /// <param name="strPool">要判斷的字詞</param>
        /// <param name="splitSymbol">Array的分割符號</param>
        /// <param name="splitNum">分割符號的數量</param>
        /// <returns></returns>
        /// <example>
        ///     string strTmp = ".jpg||.png||.pdf||.bmp";
        ///     Response.Write(fn_Extensions.CheckStrWord(src, strTmp, "|", 2));        
        /// </example>
        public static bool CheckStrWord(string inputValue, string strPool, string splitSymbol, int splitNum)
        {
            string[] strAry = Regex.Split(strPool, @"\" + splitSymbol + "{" + splitNum + "}");
            foreach (string item in strAry)
            {
                if (inputValue.IndexOf(item.ToString(), StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 產生GUID
        /// </summary>
        /// <returns></returns>
        public static string GetGuid()
        {
            return System.Guid.NewGuid().ToString();
        }


        /// <summary>
        /// 取得TimeStamp String
        /// </summary>
        /// <returns></returns>
        public static string GetTS()
        {
            long ts = Cryptograph.GetCurrentTime();
            return ts.ToString();
        }


        /// <summary>
        /// 取得 N 個工作日後的日期
        /// </summary>
        /// <param name="inputDate">輸入日期</param>
        /// <param name="inputDays">N</param>
        /// <returns></returns>
        public static DateTime GetWorkDate(DateTime inputDate, int inputDays)
        {
            DateTime tempDT = inputDate;
            while (inputDays-- > 0)
            {
                tempDT = tempDT.AddDays(1);
                while (tempDT.DayOfWeek == DayOfWeek.Saturday || tempDT.DayOfWeek == DayOfWeek.Sunday)
                {
                    tempDT = tempDT.AddDays(1);
                }
            }
            return tempDT;
        }


        /// <summary>
        /// 取得兩日期間的工作日
        /// </summary>
        /// <param name="sDate"></param>
        /// <param name="eDate"></param>
        /// <returns></returns>
        public static int GetWorkDays(DateTime sDate, DateTime eDate)
        {
            DateTime tempDT;
            int sumDays = 0;

            for (int row = 0; row < ((TimeSpan)(eDate - sDate)).Days + 1; row++)
            {
                tempDT = sDate.AddDays(row);
                if (tempDT.DayOfWeek != DayOfWeek.Saturday && tempDT.DayOfWeek != DayOfWeek.Sunday)
                {
                    sumDays++;
                }

            }

            return sumDays;
        }


        #endregion

        #region -- 字串驗証 --

        //================================= 字串 =================================
        public enum InputType
        {
            英文,
            數字,
            小寫英文,
            小寫英文混數字,
            小寫英文開頭混數字,
            大寫英文,
            大寫英文混數字,
            大寫英文開頭混數字
        }

        /// <summary>
        /// 驗証 - 輸入類型(文字)
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="InputType">輸入類型</param>
        /// <param name="minLength">最少字元數</param>
        /// <param name="maxLength">最大字元數</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool String_輸入限制(string value, InputType InputType, string minLength, string maxLength
            , out string ErrMsg)
        {
            try
            {
                value = value.Trim();
                ErrMsg = "";

                //判斷輸入限制種類 - InputType
                switch (InputType)
                {
                    case InputType.數字:

                        return IsNumeric(value);

                    case InputType.英文:
                        for (int i = 0; i < value.Length; i++)
                        {
                            if ((System.Char.Parse(value.Substring(i, 1)) < 65 | System.Char.Parse(value.Substring(i, 1)) > 90)
                                & System.Char.Parse(value.Substring(i, 1)) < 97 | System.Char.Parse(value.Substring(i, 1)) > 122)
                            {
                                return false;
                            }
                        }

                        break;

                    case InputType.小寫英文:
                        for (int i = 0; i < value.Length; i++)
                        {
                            if ((System.Char.Parse(value.Substring(i, 1)) < 97 | System.Char.Parse(value.Substring(i, 1)) > 122))
                            {
                                return false;
                            }
                        }

                        break;

                    case InputType.小寫英文混數字:
                        for (int i = 0; i < value.Length; i++)
                        {
                            if ((System.Char.Parse(value.Substring(i, 1)) < 97 | System.Char.Parse(value.Substring(i, 1)) > 122)
                                & (System.Char.Parse(value.Substring(i, 1)) < 48 | System.Char.Parse(value.Substring(i, 1)) > 57))
                            {
                                return false;
                            }
                        }

                        break;

                    case InputType.小寫英文開頭混數字:
                        for (int i = 0; i < value.Length; i++)
                        {
                            if (i == 0)
                            {
                                if ((System.Char.Parse(value.Substring(i, 1)) < 97 | System.Char.Parse(value.Substring(i, 1)) > 122))
                                {
                                    return false;
                                }
                            }
                            else
                            {
                                if ((System.Char.Parse(value.Substring(i, 1)) < 97 | System.Char.Parse(value.Substring(i, 1)) > 122)
                                    & (System.Char.Parse(value.Substring(i, 1)) < 48 | System.Char.Parse(value.Substring(i, 1)) > 57))
                                {
                                    return false;
                                }
                            }
                        }

                        break;

                    case InputType.大寫英文:
                        for (int i = 0; i < value.Length; i++)
                        {
                            if ((System.Char.Parse(value.Substring(i, 1)) < 65 | System.Char.Parse(value.Substring(i, 1)) > 90))
                            {
                                return false;
                            }
                        }

                        break;

                    case InputType.大寫英文混數字:
                        for (int i = 0; i < value.Length; i++)
                        {
                            if ((System.Char.Parse(value.Substring(i, 1)) < 65 | System.Char.Parse(value.Substring(i, 1)) > 90)
                                & (System.Char.Parse(value.Substring(i, 1)) < 48 | System.Char.Parse(value.Substring(i, 1)) > 57))
                            {
                                return false;
                            }
                        }

                        break;

                    case InputType.大寫英文開頭混數字:
                        for (int i = 0; i < value.Length; i++)
                        {
                            if (i == 0)
                            {
                                if ((System.Char.Parse(value.Substring(i, 1)) < 65 | System.Char.Parse(value.Substring(i, 1)) > 90))
                                {
                                    return false;
                                }
                            }
                            else
                            {
                                if ((System.Char.Parse(value.Substring(i, 1)) < 65 | System.Char.Parse(value.Substring(i, 1)) > 90)
                                    & (System.Char.Parse(value.Substring(i, 1)) < 48 | System.Char.Parse(value.Substring(i, 1)) > 57))
                                {
                                    return false;
                                }
                            }
                        }

                        break;
                }

                //檢查字數是不是小於 minLength
                if (IsNumeric(minLength))
                {
                    if (value.Length < Math.Floor(Convert.ToDouble(minLength)))
                    {
                        ErrMsg = "字數小於 minLength：" + Math.Floor(Convert.ToDouble(minLength));
                        return false;
                    }
                }
                //檢查字數是不是大於 maxLength
                if (IsNumeric(maxLength))
                {
                    if (value.Length > Math.Floor(Convert.ToDouble(maxLength)))
                    {
                        ErrMsg = "字數大於 maxLength：" + maxLength;
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }

        /// <summary>
        /// 驗証 - 輸入字數(文字)
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minLength">最少字元數</param>
        /// <param name="maxLength">最大字元數</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool String_字數(string value, string minLength, string maxLength, out string ErrMsg)
        {
            try
            {
                value = value.Trim();
                ErrMsg = "";

                //檢查字數是不是小於 minLength
                if (IsNumeric(minLength))
                {
                    if (value.Length < Math.Floor(Convert.ToDouble(minLength)))
                    {
                        ErrMsg = "字數小於 minLength：" + Math.Floor(Convert.ToDouble(minLength));
                        return false;
                    }
                }
                //檢查字數是不是大於 maxLength
                if (IsNumeric(maxLength))
                {
                    if (value.Length > Math.Floor(Convert.ToDouble(maxLength)))
                    {
                        ErrMsg = "字數大於 maxLength：" + maxLength;
                        return false;
                    }
                }

                return true;

            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;
            }
        }

        /// <summary>
        /// 驗証 - 輸入字數(byte)(文字)
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minLength">最少字元數</param>
        /// <param name="maxLength">最大字元數</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool String_資料長度Byte(string value, string minLength, string maxLength, out string ErrMsg)
        {
            try
            {
                value = value.Trim();
                ErrMsg = "";

                double valueByteLength = System.Text.Encoding.Default.GetBytes(value).Length;
                //檢查資料長度(Byte)是不是小於 minLength
                if (IsNumeric(minLength))
                {
                    if (valueByteLength < Math.Floor(Convert.ToDouble(minLength)))
                    {
                        ErrMsg = "資料長度(Byte)小於 minLength：" + Math.Floor(Convert.ToDouble(minLength));
                        return false;
                    }
                }
                //檢查資料長度(Byte)是不是大於 maxLength
                if (IsNumeric(maxLength))
                {
                    if (valueByteLength > Math.Floor(Convert.ToDouble(maxLength)))
                    {
                        ErrMsg = "資料長度(Byte)大於 maxLength：" + maxLength;
                        return false;
                    }
                }

                return true;

            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }

        //================================ 日期時間 ==============================
        /// <summary>
        /// 驗証 - 日期
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minDate">最小日期</param>
        /// <param name="maxDate">最大日期</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool DateTime_日期(string value, string minDate, string maxDate, out string ErrMsg)
        {
            try
            {
                DateTime dtResult;
                ErrMsg = "";
                value = value.Trim();
                minDate = minDate.Trim();
                maxDate = maxDate.Trim();
                //檢查是不是時間
                if (DateTime.TryParse(value, out dtResult) == false | string.IsNullOrEmpty(value))
                {
                    ErrMsg = "不是日期資料";
                    return false;
                }
                //檢查是不是小於 minDate
                if (DateTime.TryParse(minDate, out dtResult) & !string.IsNullOrEmpty(minDate))
                {
                    if (Convert.ToDateTime(value) < Convert.ToDateTime(minDate))
                    {
                        ErrMsg = "小於 minDate：" + string.Format(minDate, "yyyy-MM-dd");
                        return false;
                    }
                }
                //檢查是不是小於 maxDate
                if (DateTime.TryParse(maxDate, out dtResult) & !string.IsNullOrEmpty(maxDate))
                {
                    if (Convert.ToDateTime(value) > Convert.ToDateTime(maxDate))
                    {
                        ErrMsg = "大於 maxDate：" + string.Format(maxDate, "yyyy-MM-dd");
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }

        /// <summary>
        /// 驗証 - 時間
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minDateTime">最小時間</param>
        /// <param name="maxDateTime">最大時間</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool DateTime_時間(string value, string minDateTime, string maxDateTime, out string ErrMsg)
        {
            try
            {
                DateTime dtResult;
                ErrMsg = "";
                value = value.Trim();
                minDateTime = minDateTime.Trim();
                maxDateTime = maxDateTime.Trim();
                //檢查是不是時間
                if (DateTime.TryParse(value, out dtResult) == false | string.IsNullOrEmpty(value))
                {
                    ErrMsg = "不是時間資料";
                    return false;
                }
                //檢查是不是小於 minDateTime
                if (DateTime.TryParse(minDateTime, out dtResult) & !string.IsNullOrEmpty(minDateTime))
                {
                    if (Convert.ToDateTime(value) < Convert.ToDateTime(minDateTime))
                    {
                        ErrMsg = "小於 minDateTime：" + string.Format(minDateTime, "yyyy-MM-dd HH:mm:ss.fff");
                        return false;
                    }
                }
                //檢查是不是小於 maxDateTime
                if (DateTime.TryParse(maxDateTime, out dtResult) & !string.IsNullOrEmpty(maxDateTime))
                {
                    if (Convert.ToDateTime(value) > Convert.ToDateTime(maxDateTime))
                    {
                        ErrMsg = "大於 maxDateTime：" + string.Format(maxDateTime, "yyyy-MM-dd HH:mm:ss.fff");
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }

        //================================= 數值 =================================
        /// <summary>
        /// 驗証 - 數字(正整數)
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minValue">最小數值</param>
        /// <param name="maxValue">最大數值</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool Num_正整數(string value, string minValue, string maxValue, out string ErrMsg)
        {
            try
            {
                value = value.Trim();
                ErrMsg = "";

                //檢查是不是數值
                if (IsNumeric(value) == false)
                {
                    ErrMsg = "不是數值";
                    return false;
                }
                //檢查是不是大於零
                if (Convert.ToDouble(value) < 0)
                {
                    ErrMsg = "小於 0";
                    return false;
                }
                //檢查是不是整數
                if (Convert.ToDouble(value) != Math.Floor(Convert.ToDouble(value)))
                {
                    ErrMsg = "正數非正整數";
                    return false;
                }
                //檢查是不是小於 minValue
                if (IsNumeric(minValue))
                {
                    if (Convert.ToDouble(value) < Convert.ToDouble(minValue))
                    {
                        ErrMsg = "小於 minValue：" + minValue;
                        return false;
                    }
                }
                //檢查是不是大於 maxValue
                if (IsNumeric(maxValue))
                {
                    if (Convert.ToDouble(value) > Convert.ToDouble(maxValue))
                    {
                        ErrMsg = "大於 maxValue：" + maxValue;
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }

        /// <summary>
        /// 驗証 - 數字(負整數)
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minValue">最小數值</param>
        /// <param name="maxValue">最大數值</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool Num_負整數(string value, string minValue, string maxValue, out string ErrMsg)
        {
            try
            {
                value = value.Trim();
                ErrMsg = "";

                //檢查是不是數值
                if (IsNumeric(value) == false)
                {
                    ErrMsg = "不是數值";
                    return false;
                }
                //檢查是不是大於零
                if (Convert.ToDouble(value) > 0)
                {
                    ErrMsg = "大於 0";
                    return false;
                }
                //檢查是不是整數
                if (Convert.ToDouble(value) != Math.Floor(Convert.ToDouble(value)))
                {
                    ErrMsg = "負數非負整數";
                    return false;
                }
                //檢查是不是小於 minValue
                if (IsNumeric(minValue))
                {
                    if (Convert.ToDouble(value) < Convert.ToDouble(minValue))
                    {
                        ErrMsg = "小於 minValue：" + minValue;
                        return false;
                    }
                }
                //檢查是不是大於 maxValue
                if (IsNumeric(maxValue))
                {
                    if (Convert.ToDouble(value) > Convert.ToDouble(maxValue))
                    {
                        ErrMsg = "大於 maxValue：" + maxValue;
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }

        /// <summary>
        /// 驗証 - 數字(正數)
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minValue">最小數值</param>
        /// <param name="maxValue">最大數值</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool Num_正數(string value, string minValue, string maxValue, out string ErrMsg)
        {
            try
            {
                value = value.Trim();
                ErrMsg = "";

                //檢查是不是數值
                if (IsNumeric(value) == false)
                {
                    ErrMsg = "不是數值";
                    return false;
                }
                //檢查是不是大於零
                if (Convert.ToDouble(value) < 0)
                {
                    ErrMsg = "小於 0";
                    return false;
                }
                //檢查是不是小於 minValue
                if (IsNumeric(minValue))
                {
                    if (Convert.ToDouble(value) < Convert.ToDouble(minValue))
                    {
                        ErrMsg = "小於 minValue：" + minValue;
                        return false;
                    }
                }
                //檢查是不是大於 maxValue
                if (IsNumeric(maxValue))
                {
                    if (Convert.ToDouble(value) > Convert.ToDouble(maxValue))
                    {
                        ErrMsg = "大於 maxValue：" + maxValue;
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }

        /// <summary>
        /// 驗証 - 數字(負數)
        /// </summary>
        /// <param name="value">要驗証的值</param>
        /// <param name="minValue">最小數值</param>
        /// <param name="maxValue">最大數值</param>
        /// <param name="ErrMsg">錯誤訊息</param>
        /// <returns>Boolean</returns>
        public static bool Num_負數(string value, string minValue, string maxValue, out string ErrMsg)
        {
            try
            {
                value = value.Trim();
                ErrMsg = "";

                //檢查是不是數值
                if (IsNumeric(value) == false)
                {
                    ErrMsg = "不是數值";
                    return false;
                }
                //檢查是不是大於零
                if (Convert.ToDouble(value) > 0)
                {
                    ErrMsg = "大於 0";
                    return false;
                }
                //檢查是不是小於 minValue
                if (IsNumeric(minValue))
                {
                    if (Convert.ToDouble(value) < Convert.ToDouble(minValue))
                    {
                        ErrMsg = "小於 minValue：" + minValue;
                        return false;
                    }
                }
                //檢查是不是大於 maxValue
                if (IsNumeric(maxValue))
                {
                    if (Convert.ToDouble(value) > Convert.ToDouble(maxValue))
                    {
                        ErrMsg = "大於 maxValue：" + maxValue;
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrMsg = "Exception：" + ex.Message.ToString();
                return false;

            }
        }


        #endregion

        #region -- 常用功能 --
        /// <summary>
        /// 使用HttpWebRequest取得網頁資料 (AD驗證模式使用)
        /// </summary>
        /// <param name="url">網址</param>
        /// <returns>string</returns>
        public static string WebRequest_GET(string url, bool ADMode)
        {
            try
            {
                Encoding myEncoding = Encoding.GetEncoding("UTF-8");
                HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

                //安全通訊協定
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls |
                    SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                //IIS為AD驗証時加入此段 Start
                if (ADMode)
                {
                    req.UseDefaultCredentials = true;
                    req.PreAuthenticate = true;
                    req.Credentials = CredentialCache.DefaultCredentials;
                }
                //IIS為AD驗証時加入此段 End

                req.Method = "GET";
                using (WebResponse wr = req.GetResponse())
                {
                    using (StreamReader myStreamReader = new StreamReader(wr.GetResponseStream(), myEncoding))
                    {
                        return myStreamReader.ReadToEnd();
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

        }

        /// <summary>
        /// 使用HttpWebRequest POST取得網頁資料 (AD驗證模式使用)
        /// </summary>
        /// <param name="isAD">是否為AD</param>
        /// <param name="url">網址</param>
        /// <param name="postParameters">參數 (a=123&b=456)</param>
        /// <param name="postHeaders">header</param>
        /// <returns></returns>
        public static string WebRequest_POST(bool isAD, string url, Dictionary<string, string> postParameters, Dictionary<string, string> postHeaders)
        {
            try
            {
                //取得傳遞參數
                string postData = "";

                if (postParameters != null)
                {
                    foreach (string key in postParameters.Keys)
                    {
                        postData += key + "="
                              + postParameters[key] + "&";
                    }
                }

                //傳遞參數轉為byte
                byte[] bs = Encoding.ASCII.GetBytes(postData);

                //設定UTF8
                Encoding myEncoding = Encoding.GetEncoding("UTF-8");

                //安全通訊協定
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls |
                    SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                //設定網址
                HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

                //IIS為AD驗証時加入此段 Start
                if (isAD)
                {
                    req.UseDefaultCredentials = true;
                    req.PreAuthenticate = true;
                    req.Credentials = CredentialCache.DefaultCredentials;
                }
                //IIS為AD驗証時加入此段 End


                req.Method = "POST";
                req.ContentType = "application/x-www-form-urlencoded";
                req.ContentLength = bs.Length;

                //自訂headers
                if (postHeaders != null)
                {
                    foreach (KeyValuePair<string, string> item in postHeaders)
                    {
                        req.Headers.Add(item.Key, item.Value);
                    }
                }

                // 寫入 Post Body Message 資料流
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(bs, 0, bs.Length);
                }

                // 取得回應資料
                using (WebResponse wr = req.GetResponse())
                {
                    using (StreamReader myStreamReader = new StreamReader(wr.GetResponseStream(), myEncoding))
                    {
                        return myStreamReader.ReadToEnd();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
        }


        /// <summary>
        /// 使用HttpClient取得遠端資料
        /// Http Method = GET
        /// </summary>
        /// <param name="url">網址</param>
        /// <returns>string</returns>
        public static string WebRequest_byGET(string url)
        {
            try
            {
                string resultData = "";

                //安全通訊協定
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls |
                    SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                //宣告HttpClient
                using (HttpClient client = new HttpClient())
                {
                    //取得Http回應訊息
                    using (HttpResponseMessage response = client.GetAsync(url).Result)
                    {
                        //Throws an exception if the HTTP response is false.
                        response.EnsureSuccessStatusCode();

                        //取得回應結果
                        resultData = response.Content.ReadAsStringAsync().Result;
                    }
                }

                //回傳資料
                return resultData;

            }
            catch (Exception)
            {

                return "error";
            }
        }


        /// <summary>
        /// 取得遠端資料 [舊版]
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postParams"></param>
        /// <returns></returns>
        public static string WebRequest_byPOST(string url, string postParams)
        {
            try
            {
                //安全通訊協定
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls |
                    SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                //param
                byte[] postData = Encoding.ASCII.GetBytes(postParams);

                Encoding myEncoding = Encoding.GetEncoding("UTF-8");
                HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

                req.Method = "POST";
                req.ContentType = "application/x-www-form-urlencoded";
                req.ContentLength = postData.Length;

                // 寫入 Post Body Message 資料流
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(postData, 0, postData.Length);
                }

                // 取得回應資料
                string result = "";
                using (HttpWebResponse wr = req.GetResponse() as HttpWebResponse)
                {
                    using (StreamReader sr = new StreamReader(wr.GetResponseStream(), myEncoding))
                    {
                        result = sr.ReadToEnd();
                    }
                }

                return result;

            }
            catch (Exception)
            {

                throw;
            }
        }



        /// <summary>
        /// 使用HttpClient取得遠端資料
        /// Http Method = POST
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postParams"></param>
        /// <param name="postHeaders"></param>
        /// <returns></returns>
        public static string WebRequest_byPOST(string url
            , Dictionary<string, string> postParams, Dictionary<string, string> postHeaders)
        {
            try
            {
                string resultData = "";

                //安全通訊協定
                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls |
                    SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                //宣告HttpClient
                using (HttpClient client = new HttpClient())
                {
                    //自訂傳遞參數(使用FormUrlEncodedContent)
                    var postData = new FormUrlEncodedContent(postParams);
                    /* 
                     * 若要使用別的格式Post資料(如json), 需使用new StringContent 如下:
                     * HttpResponseMessage wcfResponse = await httpClient.PostAsync(resourceAddress, new StringContent(postBody, Encoding.UTF8, "application/json")); 
                     */

                    //自訂headers
                    /*
                     * 若要使用別的標頭傳遞Post資料(如json), 需定義header 如下:
                     * httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json")); 
                     * 
                     * 或是使用 PostAsJsonAsync(未測)
                     */
                    if (postHeaders != null)
                    {
                        foreach (KeyValuePair<string, string> item in postHeaders)
                        {
                            client.DefaultRequestHeaders.Add(item.Key, item.Value);
                        }
                    }

                    //取得Http回應訊息
                    using (HttpResponseMessage response = client.PostAsync(url, postData).Result)
                    {
                        //Throws an exception if the HTTP response is false.
                        response.EnsureSuccessStatusCode();

                        //取得回應結果
                        resultData = response.Content.ReadAsStringAsync().Result;

                    }
                }

                //回傳資料
                return resultData;

            }
            catch (Exception)
            {

                throw;
            }
        }


        /// <summary>
        /// 使用HttpClient取得遠端資料
        /// Http Method = POST
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postParams"></param>
        /// <param name="postHeaders"></param>
        /// <returns>byte[]</returns>
        public static byte[] WebRequestByte_byPOST(string url
            , Dictionary<string, string> postParams, Dictionary<string, string> postHeaders)
        {
            try
            {
                byte[] resultData = null;

                //宣告HttpClient
                using (HttpClient client = new HttpClient())
                {
                    //安全通訊協定
                    ServicePointManager.SecurityProtocol =
                        SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls |
                        SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;


                    //自訂傳遞參數(使用FormUrlEncodedContent)
                    var postData = new FormUrlEncodedContent(postParams);
                    /* 
                     * 若要使用別的格式Post資料(如json), 需使用new StringContent 如下:
                     * HttpResponseMessage wcfResponse = await httpClient.PostAsync(resourceAddress, new StringContent(postBody, Encoding.UTF8, "application/json")); 
                     */

                    //自訂headers
                    /*
                     * 若要使用別的標頭傳遞Post資料(如json), 需定義header 如下:
                     * httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json")); 
                     * 
                     * 或是使用 PostAsJsonAsync(未測)
                     */
                    if (postHeaders != null)
                    {
                        foreach (KeyValuePair<string, string> item in postHeaders)
                        {
                            client.DefaultRequestHeaders.Add(item.Key, item.Value);
                        }
                    }

                    //取得Http回應訊息
                    using (HttpResponseMessage response = client.PostAsync(url, postData).Result)
                    {
                        //Throws an exception if the HTTP response is false.
                        response.EnsureSuccessStatusCode();

                        //取得回應結果
                        resultData = response.Content.ReadAsByteArrayAsync().Result; ;

                    }
                }

                //回傳資料
                return resultData;

            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// 使用FileStream取得資料
        /// </summary>
        /// <param name="path">磁碟路徑</param>
        /// <returns>string</returns>
        public static string IORequest_GET(string path)
        {
            try
            {
                if (false == System.IO.File.Exists(path)) return "";
                using (FileStream fs = new FileStream(path, FileMode.Open))
                {
                    using (StreamReader sr = new StreamReader(fs, System.Text.Encoding.UTF8))
                    {
                        return sr.ReadToEnd();
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }


        /// <summary>
        /// 顯示Alert
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="url"></param>
        public static void AlertMsg(string msg, string url)
        {
            StringBuilder sbJs = new StringBuilder();

            //Message
            sbJs.Append("alert('{0}');".FormatThis(msg));

            //Url
            if (!string.IsNullOrEmpty(url))
            {
                sbJs.Append("location.href='{0}';".FormatThis(url));
            }

            System.Web.UI.ScriptManager.RegisterClientScriptBlock((System.Web.UI.Page)HttpContext.Current.Handler, typeof(string), "js", sbJs.ToString(), true);
        }


        #endregion

        #region -- 產生分頁 --

        public enum myStyle : int
        {
            Goole = 1,
            Bootstrap = 2
        }

        public static string PageControl(int TotalRow, int PageSize, int CurrentPageIdx, int PageRoll, string PageUrl, ArrayList Params, bool IsRouting)
        {
            return PageControl(TotalRow, PageSize, CurrentPageIdx, PageRoll, PageUrl, Params, IsRouting, true, myStyle.Goole);
        }

        /// <summary>
        /// 自訂分頁
        /// </summary>
        /// <param name="TotalRow">總筆數</param>
        /// <param name="PageSize">每頁顯示筆數</param>
        /// <param name="CurrentPageIdx">目前的索引頁</param>
        /// <param name="PageRoll">要顯示幾個頁碼</param>
        /// <param name="PageUrl">Url</param>
        /// <param name="Params">參數</param>
        /// <param name="IsRouting">是否使用Routing</param>
        /// <param name="showInfo"></param>
        /// <param name="style"></param>
        /// <returns>string</returns>
        public static string PageControl(int TotalRow, int PageSize, int CurrentPageIdx, int PageRoll, string PageUrl, ArrayList Params, bool IsRouting
            , bool showInfo, myStyle style)
        {
            //[參數宣告]
            int cntBgNum, cntEdNum;     //計算開始數, 計算終止數
            int PageBg, PageEd;     //設定頁數(開始), 設定頁數(結束)

            //[參數設定] - 計算總頁數, TotalPage
            int TotalPage = (TotalRow / PageSize);
            if (TotalRow % PageSize > 0)
            {
                TotalPage++;
            }
            //[參數設定] - 判斷Request Page, 若目前Page < 1, Page設為 1
            if (CurrentPageIdx < 1)
            {
                CurrentPageIdx = 1;
            }
            //[參數設定] - 判斷Request Page, 若目前Page > 總頁數TotalPage, Page 設為 TotalPage
            if (CurrentPageIdx > TotalPage)
            {
                CurrentPageIdx = TotalPage;
            }
            //[參數設定] - 開始資料列/結束資料列 (分頁資訊)
            int FirstItem = (CurrentPageIdx - 1) * PageSize + 1;
            int LastItem = FirstItem + (PageSize - 1);
            if (LastItem > TotalRow)
            {
                LastItem = TotalRow;
            }

            //頁數資訊
            string PageInfo = string.Format("顯示第 {0} - {1} 筆,共 {2} 筆", FirstItem, LastItem, TotalRow);

            //[分頁設定] - 計算開始頁/結束頁
            cntBgNum = CurrentPageIdx - ((PageRoll + 5) / 5);
            cntEdNum = CurrentPageIdx + ((PageRoll + 5) / 5);

            //[分頁設定] - 設定開始值/結束值
            PageBg = cntBgNum;
            PageEd = cntEdNum;

            //判斷開始值 是否小於 1
            if (PageBg < 1)
            {
                PageBg = 1;
                PageEd = (cntEdNum - cntBgNum) + 1;
            }
            //判斷結束值 是否大於 總頁數
            if (PageEd > TotalPage)
            {
                if (cntBgNum > 1)
                {
                    PageBg = cntBgNum - (cntEdNum - TotalPage);
                    if (PageBg == 0) PageBg = 1;
                }
                PageEd = TotalPage;
            }

            //----- 分頁Html -----
            StringBuilder sb = new StringBuilder();

            if (showInfo)
            {
                sb.AppendLine("<div class=\"{1}\">{0}</div>".FormatThis(
                    PageInfo
                    , style.Equals(myStyle.Goole) ? "left-align" : "text-left"));
            }
            sb.AppendLine("<div class=\"{0}\">".FormatThis(style.Equals(myStyle.Goole) ? "right-align" : "text-right"));

            sb.AppendLine("<ul class=\"pagination\">");

            string fixParams = "";
            //判斷參數串
            if (Params != null && Params.Count > 0)
            {
                if (IsRouting)
                {
                    fixParams = "?" + string.Join("&", Params.ToArray());
                }
                else
                {
                    fixParams = "&" + string.Join("&", Params.ToArray());
                }
            }

            //[分頁按鈕] - 第一頁 & 上一頁
            if (CurrentPageIdx > 1)
            {
                sb.Append("<li>");

                //第一頁
                //sb.AppendFormat("<a href=\"{0}{1}{2}\"><i class=\"material-icons\">first_page</i></a> ", PageUrl
                //    , (IsRouting) ? "/{0}/".FormatThis(1) : "?page=1"
                //    , fixParams);

                //上一頁
                sb.AppendFormat("<a href=\"{0}{1}{2}\"><span>{3}</span></a> ", PageUrl
                    , (IsRouting) ? "/{0}/".FormatThis(CurrentPageIdx - 1) : "?page={0}".FormatThis(CurrentPageIdx - 1)
                    , fixParams
                    , style.Equals(myStyle.Goole) ? "<i class=\"material-icons\">chevron_left</i>" : "&larr;");

                sb.Append("</li>");
            }
            else
            {
                //上一頁 - disabled
                if (style.Equals(myStyle.Goole))
                {
                    sb.AppendLine("<li class=\"disabled\"><a href=\"#!\"><i class=\"material-icons\">chevron_left</i></a></li>");
                }
                else
                {
                    sb.AppendLine("<li class=\"disabled\"><a>&larr;</a></li>");
                }
            }

            //[分頁按鈕] - 數字頁碼
            for (int row = PageBg; row <= PageEd; row++)
            {
                if (row == CurrentPageIdx)
                {
                    sb.Append("<li class=\"active\">");
                    sb.AppendFormat("<a>{0}</a>", row);
                    sb.Append("</li>");
                }
                else
                {
                    sb.Append("<li>");
                    sb.AppendFormat("<a href=\"{0}{1}{2}\">{3}</a> ", PageUrl
                    , (IsRouting) ? "/{0}/".FormatThis(row) : "?page={0}".FormatThis(row)
                    , fixParams
                    , row);
                    sb.Append("</li>");
                }
            }

            //[分頁按鈕] - 最後一頁 & 下一頁
            if (CurrentPageIdx < TotalPage)
            {
                sb.Append("<li>");

                //下一頁
                sb.AppendFormat("<a href=\"{0}{1}{2}\"><span>{3}</span></a> ", PageUrl
                    , (IsRouting) ? "/{0}/".FormatThis(CurrentPageIdx + 1) : "?page={0}".FormatThis(CurrentPageIdx + 1)
                    , fixParams
                    , style.Equals(myStyle.Goole) ? "<i class=\"material-icons\">chevron_right</i>" : "&rarr;");

                //最後一頁
                //sb.AppendFormat("<a href=\"{0}{1}{2}\"><span>{3}</span></a> ", PageUrl
                //    , (IsRouting) ? "/{0}/".FormatThis(TotalPage) : "?page={0}".FormatThis(TotalPage)
                //    , fixParams
                //    , "<i class=\"material-icons\">last_page</i>");


                sb.Append("</li>");
            }
            else
            {
                //下一頁 - disabled
                if (style.Equals(myStyle.Goole))
                {
                    sb.AppendLine("<li class=\"disabled\"><a href=\"#!\"><i class=\"material-icons\">chevron_right</i></a></li>");
                }
                else
                {
                    sb.AppendLine("<li class=\"disabled\"><a>Next &rarr;</a></li>");
                }
            }


            sb.AppendLine("</ul>");
            sb.AppendLine("</div>");

            //[輸出Html]
            return sb.ToString();
        }


        /// <summary>
        /// 自訂分頁  SemanticUI
        /// </summary>
        /// <param name="TotalRow">總筆數</param>
        /// <param name="PageSize">每頁顯示筆數</param>
        /// <param name="CurrentPageIdx">目前的索引頁</param>
        /// <param name="PageRoll">要顯示幾個頁碼</param>
        /// <param name="PageUrl">Url</param>
        /// <param name="Params">參數</param>
        /// <param name="IsRouting">是否使用Routing</param>
        /// <param name="showInfo">是否顯示頁數資訊</param>
        /// <returns>string</returns>
        public static string Pagination(int TotalRow, int PageSize, int CurrentPageIdx, int PageRoll, string PageUrl
            , ArrayList Params, bool IsRouting, bool showInfo)
        {
            //[參數宣告]
            int cntBgNum, cntEdNum;     //計算開始數, 計算終止數
            int PageBg, PageEd;     //設定頁數(開始), 設定頁數(結束)

            //[參數設定] - 計算總頁數, TotalPage
            int TotalPage = (TotalRow / PageSize);
            if (TotalRow % PageSize > 0)
            {
                TotalPage++;
            }
            //[參數設定] - 判斷Request Page, 若目前Page < 1, Page設為 1
            if (CurrentPageIdx < 1)
            {
                CurrentPageIdx = 1;
            }
            //[參數設定] - 判斷Request Page, 若目前Page > 總頁數TotalPage, Page 設為 TotalPage
            if (CurrentPageIdx > TotalPage)
            {
                CurrentPageIdx = TotalPage;
            }
            //[參數設定] - 開始資料列/結束資料列 (分頁資訊)
            int FirstItem = (CurrentPageIdx - 1) * PageSize + 1;
            int LastItem = FirstItem + (PageSize - 1);
            if (LastItem > TotalRow)
            {
                LastItem = TotalRow;
            }

            //頁數資訊
            string PageInfo = string.Format("顯示第&nbsp;<b>{0} - {1}</b>&nbsp;筆,&nbsp;共&nbsp;<b>{2}</b>&nbsp;筆", FirstItem, LastItem, TotalRow);

            //[分頁設定] - 計算開始頁/結束頁
            cntBgNum = CurrentPageIdx - ((PageRoll + 5) / 5);
            cntEdNum = CurrentPageIdx + ((PageRoll + 5) / 5);

            //[分頁設定] - 設定開始值/結束值
            PageBg = cntBgNum;
            PageEd = cntEdNum;

            //判斷開始值 是否小於 1
            if (PageBg < 1)
            {
                PageBg = 1;
                PageEd = (cntEdNum - cntBgNum) + 1;
            }
            //判斷結束值 是否大於 總頁數
            if (PageEd > TotalPage)
            {
                if (cntBgNum > 1)
                {
                    PageBg = cntBgNum - (cntEdNum - TotalPage);
                    if (PageBg == 0) PageBg = 1;
                }
                PageEd = TotalPage;
            }

            //----- 分頁Html -----
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("<div class=\"ui two column grid\">");

            if (showInfo)
            {
                //Page info
                sb.AppendLine("<div class=\"column\">{0}</div>".FormatThis(PageInfo));
            }

            sb.AppendLine("<div class=\"column right aligned\"><div class=\"ui small red pagination menu\">");

            string fixParams = "";
            //判斷參數串
            if (Params != null && Params.Count > 0)
            {
                if (IsRouting)
                {
                    fixParams = "?" + string.Join("&", Params.ToArray());
                }
                else
                {
                    fixParams = "&" + string.Join("&", Params.ToArray());
                }
            }

            //[分頁按鈕] - 第一頁 & 上一頁
            if (CurrentPageIdx > 1)
            {
                //第一頁
                //sb.AppendFormat("<a href=\"{0}{1}{2}\"><i class=\"material-icons\">first_page</i></a> ", PageUrl
                //    , (IsRouting) ? "/{0}/".FormatThis(1) : "?page=1"
                //    , fixParams);

                //上一頁
                sb.AppendFormat("<a class=\"icon item\" href=\"{0}{1}{2}\">{3}</a> "
                    , PageUrl
                    , (IsRouting) ? "/{0}/".FormatThis(CurrentPageIdx - 1) : "?page={0}".FormatThis(CurrentPageIdx - 1)
                    , fixParams
                    , "<i class=\"left chevron icon\"></i>");
            }
            else
            {
                //上一頁 - disabled
                sb.AppendLine("<div class=\"disabled item\"><i class=\"left chevron icon\"></i></div>");
            }

            //[分頁按鈕] - 數字頁碼
            for (int row = PageBg; row <= PageEd; row++)
            {
                if (row == CurrentPageIdx)
                {
                    sb.Append("<a class=\"active item\"><b>{0}</b></a>".FormatThis(row));
                }
                else
                {
                    sb.Append("<a class=\"item\" href=\"{0}{1}{2}\">{3}</a>".FormatThis(
                        PageUrl
                        , (IsRouting) ? "/{0}/".FormatThis(row) : "?page={0}".FormatThis(row)
                        , fixParams
                        , row
                        ));
                }
            }

            //[分頁按鈕] - 最後一頁 & 下一頁
            if (CurrentPageIdx < TotalPage)
            {
                //下一頁
                sb.AppendFormat("<a class=\"icon item\" href=\"{0}{1}{2}\">{3}</a> "
                    , PageUrl
                    , (IsRouting) ? "/{0}/".FormatThis(CurrentPageIdx + 1) : "?page={0}".FormatThis(CurrentPageIdx + 1)
                    , fixParams
                    , "<i class=\"right chevron icon\"></i>");

                //最後一頁
                //sb.AppendFormat("<a href=\"{0}{1}{2}\"><span>{3}</span></a> ", PageUrl
                //    , (IsRouting) ? "/{0}/".FormatThis(TotalPage) : "?page={0}".FormatThis(TotalPage)
                //    , fixParams
                //    , "<i class=\"material-icons\">last_page</i>");

            }
            else
            {
                //下一頁 - disabled
                sb.AppendLine("<div class=\"disabled item\"><i class=\"right chevron icon\"></i></div>");
            }

            sb.AppendLine("</div></div></div>");

            //[輸出Html]
            return sb.ToString();
        }

        #endregion

        #region -- EXCEL匯出 --

        /// <summary>
        /// 匯出Excel, 預設鎖密碼
        /// </summary>
        /// <param name="DT">DataTable</param>
        /// <param name="fileName">匯出檔名</param>
        public static void ExportExcel(DataTable DT, string fileName)
        {
            //default
            ExportExcel(DT, fileName, true);
        }

        /// <summary>
        /// 匯出Excel
        /// </summary>
        /// <param name="DT">DataTable</param>
        /// <param name="fileName">匯出檔名</param>
        /// <param name="setPassword">true/false</param>
        /// <remarks>
        /// 使用元件:ClosedXML
        /// </remarks>
        /// <seealso cref="https://github.com/ClosedXML/ClosedXML/wiki"/>
        public static void ExportExcel(DataTable DT, string fileName, bool setPassword)
        {
            ExportExcel(DT, fileName, setPassword, "PKDataList");
        }

        /// <summary>
        /// 匯出Excel
        /// </summary>
        /// <param name="DT">DataTable</param>
        /// <param name="fileName">匯出檔名</param>
        /// <param name="setPassword">true/false</param>
        /// <param name="sheetName">工作表名稱</param>
        public static void ExportExcel(DataTable DT, string fileName, bool setPassword, string sheetName)
        {
            //宣告
            XLWorkbook wbook = new XLWorkbook();

            //-- 工作表設定 Start --
            var ws = wbook.Worksheets.Add(DT, sheetName);

            if (setPassword)
            {
                //鎖定工作表, 並設定密碼
                ws.Protect("iLoveProkits25")    //Set Password
                    .SetFormatCells(true)   // Cell Formatting
                    .SetInsertColumns() // Inserting Columns
                    .SetDeleteColumns() // Deleting Columns
                    .SetDeleteRows();   // Deleting Rows
            }

            //細項設定
            ws.Tables.FirstOrDefault().ShowAutoFilter = false;  //停用自動篩選
            ws.Style.Font.FontName = "Microsoft JhengHei";  //字型名稱
            ws.Style.Font.FontSize = 10;

            //修改標題列
            var header = ws.FirstRowUsed(false);
            //header.Style.Fill.BackgroundColor = XLColor.Green;
            //header.Style.Font.FontColor = XLColor.Yellow;
            header.Style.Font.FontSize = 12;
            header.Style.Font.Bold = true;
            header.Height = 22;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //-- 工作表設定 End --

            //Http Response & Request
            var resp = HttpContext.Current.Response;
            var req = HttpContext.Current.Request;
            HttpResponse httpResponse = resp;
            httpResponse.Clear();
            // 編碼
            httpResponse.ContentEncoding = Encoding.UTF8;
            // 設定網頁ContentType
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            // 匯出檔名
            var browser = req.Browser.Browser;
            var exportFileName = browser.Equals("Firefox", StringComparison.OrdinalIgnoreCase)
                ? fileName
                : HttpUtility.UrlEncode(fileName, Encoding.UTF8);

            resp.AddHeader(
                "Content-Disposition",
                string.Format("attachment;filename={0}", exportFileName));

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
                //memoryStream.ToArray(); 轉成byte
            }

            httpResponse.End();
        }

        /// <summary>
        /// 匯出Excel, 多張工作表
        /// </summary>
        /// <param name="dataList">列舉 DataTable</param>
        /// <param name="fileName">匯出檔名</param>
        /// <param name="setPassword">true/false</param>
        /// <param name="sheetList">列舉 string</param>
        public static void ExportExcelMultiple(List<DataTable> dataList, string fileName, bool setPassword, List<string> sheetList)
        {
            //宣告
            XLWorkbook wbook = new XLWorkbook();

            //-- 工作表設定 Start --
            for (int row = 0; row < dataList.Count; row++)
            {
                DataTable DT = dataList[row];
                string sheetName = sheetList[row];

                var ws = wbook.Worksheets.Add(DT, sheetName);

                if (setPassword)
                {
                    //鎖定工作表, 並設定密碼
                    ws.Protect("iLoveProkits25")    //Set Password
                        .SetFormatCells(true)   // Cell Formatting
                        .SetInsertColumns() // Inserting Columns
                        .SetDeleteColumns() // Deleting Columns
                        .SetDeleteRows();   // Deleting Rows
                }

                //細項設定
                ws.Tables.FirstOrDefault().ShowAutoFilter = false;  //停用自動篩選
                ws.Style.Font.FontName = "Microsoft JhengHei";  //字型名稱
                ws.Style.Font.FontSize = 10;

                //修改標題列
                var header = ws.FirstRowUsed(false);
                //header.Style.Fill.BackgroundColor = XLColor.Green;
                //header.Style.Font.FontColor = XLColor.Yellow;
                header.Style.Font.FontSize = 12;
                header.Style.Font.Bold = true;
                header.Height = 22;
                header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }

            //-- 工作表設定 End --

            //Http Response & Request
            var resp = HttpContext.Current.Response;
            var req = HttpContext.Current.Request;
            HttpResponse httpResponse = resp;
            httpResponse.Clear();
            // 編碼
            httpResponse.ContentEncoding = Encoding.UTF8;
            // 設定網頁ContentType
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            // 匯出檔名
            var browser = req.Browser.Browser;
            var exportFileName = browser.Equals("Firefox", StringComparison.OrdinalIgnoreCase)
                ? fileName
                : HttpUtility.UrlEncode(fileName, Encoding.UTF8);

            resp.AddHeader(
                "Content-Disposition",
                string.Format("attachment;filename={0}", exportFileName));

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wbook.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
                //memoryStream.ToArray(); 轉成byte
            }

            httpResponse.End();
        }



        /// <summary>
        /// 產生Excel - byte[]
        /// </summary>
        /// <param name="DT">DataTable</param>
        /// <param name="setPassword">password</param>
        /// <returns>byte</returns>
        public static byte[] ExcelToByte(DataTable DT, bool setPassword)
        {
            //宣告
            XLWorkbook wbook = new XLWorkbook();

            //-- 工作表設定 Start --
            var ws = wbook.Worksheets.Add(DT, "PKDataList");

            if (setPassword)
            {
                //鎖定工作表, 並設定密碼
                ws.Protect("iLoveProkits25")    //Set Password
                    .SetFormatCells(true)   // Cell Formatting
                    .SetInsertColumns() // Inserting Columns
                    .SetDeleteColumns() // Deleting Columns
                    .SetDeleteRows();   // Deleting Rows
            }

            //細項設定
            ws.Tables.FirstOrDefault().ShowAutoFilter = false;  //停用自動篩選
                                                                //ws.Style.Font.FontName = "Microsoft JhengHei";  //字型名稱
                                                                //ws.Style.Font.FontSize = 10;

            //修改標題列
            var header = ws.FirstRowUsed(false);
            //header.Style.Fill.BackgroundColor = XLColor.Green;
            //header.Style.Font.FontColor = XLColor.Yellow;
            header.Style.Font.FontSize = 12;
            header.Style.Font.Bold = true;
            header.Height = 22;
            header.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            //-- 工作表設定 End --

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                wbook.SaveAs(memoryStream);

                byte[] ms = memoryStream.ToArray(); //轉成byte
                memoryStream.Close();

                return ms;
            }

        }


        /// <summary>
        /// Linq查詢結果轉Datatable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="query"></param>
        /// <returns></returns>
        /// <remarks>
        /// 此方法僅可接受IEnumerable<T>泛型物件
        /// DataTable dt = LINQToDataTable(query);
        /// </remarks>
        public static DataTable LINQToDataTable<T>(IEnumerable<T> query)
        {
            //宣告一個datatable
            DataTable tbl = new DataTable();
            //宣告一個propertyinfo為陣列的物件，此物件需要import reflection才可以使用
            //使用 ParameterInfo 的執行個體來取得有關參數的資料型別、預設值等資訊

            PropertyInfo[] props = null;
            //使用型別為T的item物件跑query的內容
            foreach (T item in query)
            {
                if (props == null) //尚未初始化
                {
                    //宣告一型別為T的t物件接收item.GetType()所回傳的物件
                    Type t = item.GetType();
                    //props接收t.GetProperties();所回傳型別為props的陣列物件
                    props = t.GetProperties();
                    //使用propertyinfo物件針對propertyinfo陣列的物件跑迴圈
                    foreach (PropertyInfo pi in props)
                    {
                        //將pi.PropertyType所回傳的物件指給型別為Type的coltype物件
                        Type colType = pi.PropertyType;
                        //針對Nullable<>特別處理
                        if (colType.IsGenericType
                            && colType.GetGenericTypeDefinition() == typeof(Nullable<>))
                            colType = colType.GetGenericArguments()[0];
                        //建立欄位
                        tbl.Columns.Add(pi.Name, colType);
                    }
                }
                //宣告一個datarow物件
                DataRow row = tbl.NewRow();
                //同樣利用PropertyInfo跑迴圈取得props的內容，並將內容放進剛所宣告的datarow中
                //接著在將該datarow加到datatable (tb1) 當中
                foreach (PropertyInfo pi in props)
                    row[pi.Name] = pi.GetValue(item, null) ?? DBNull.Value;
                tbl.Rows.Add(row);
            }
            //回傳tb1的datatable物件
            return tbl;
        }


        /// <summary>
        /// 排除使用者KEY在輸入欄裡特殊字元
        /// </summary>
        /// <param name="tmp">
        /// <returns></returns>
        /// <remarks>
        /// ex:(十六進位值 0x0B) 是無效的字元
        /// </remarks>
        public static string ReplaceLowOrderASCIICharacters(string tmp)
        {
            StringBuilder info = new StringBuilder();
            foreach (char cc in tmp)
            {
                int ss = (int)cc;
                if (((ss >= 0) && (ss <= 8)) || ((ss >= 11) && (ss <= 12)) || ((ss >= 14) && (ss <= 32)))
                    info.AppendFormat(" ", ss);//&#x{0:X};
                else info.Append(cc);
            }
            return info.ToString();
        }

        #endregion

        #region -- SQL功能 --
        /// <summary>
        /// SQL參數組合 - Where IN
        /// </summary>
        /// <param name="ary">來源資料</param>
        /// <param name="paramName">參數名稱</param>
        /// <returns>參數字串</returns>
        public static string GetSQLParam(ArrayList ary, string paramName)
        {
            /* example
                sql.AppendLine(" AND RTRIM(Model_No) IN ({0})".FormatThis(GetSQLParam(ary, "params")));
                for (int row = 0; row < ary.Count; row++)
                {
                    cmd.Parameters.AddWithValue("params" + row, ary[row]);
                }
            */

            if (ary.Count == 0)
            {
                return "";
            }

            //組合參數字串
            ArrayList aryParam = new ArrayList();
            for (int row = 0; row < ary.Count; row++)
            {
                aryParam.Add(string.Format("@{0}{1}", paramName, row));
            }

            //回傳以 , 為分隔符號的字串
            return string.Join(",", aryParam.ToArray());
        }
        #endregion

        #region -- Cookies功能 --
        /// <summary>
        /// 設定Cookies
        /// </summary>
        /// <param name="ckName">名稱</param>
        /// <param name="ckValue">傳入值</param>
        /// <param name="expireHours">小時</param>
        /// <returns></returns>
        /// <example>
        /// setCookie("Name", "helloValue", 1);
        /// </example>
        public static void setCookie(string ckName, string ckValue, int expireHours)
        {
            //取得目前cookie
            var requestCookie = HttpContext.Current.Request.Cookies[ckName];

            //判斷cookie是否存在
            if (requestCookie != null)
            {
                //cookie存在, 判斷內容與新設定值是否相同
                if (!requestCookie.Value.Equals(ckValue))
                {
                    //Reset Cookie
                    resetCookie(ckName, ckValue, expireHours);
                }
            }
            else
            {
                //Reset Cookie
                resetCookie(ckName, ckValue, expireHours);
            }
        }

        /// <summary>
        /// 取得Cookies
        /// </summary>
        /// <param name="ckName">名稱</param>
        /// <returns></returns>
        /// <example>
        /// string val = getCookie("Name");
        /// </example>
        public static string getCookie(string ckName)
        {
            //Get Cookie Value
            var respCookie = HttpContext.Current.Request.Cookies[ckName];

            if (respCookie == null)
            {
                return "";
            }
            else
            {
                return respCookie.Value;
            }
        }

        private static void resetCookie(string ckName, string ckValue, int expireHours)
        {
            // 產生新的值並儲存到 cookie
            var responseCookie = new System.Web.HttpCookie(ckName)
            {
                HttpOnly = true,
                Value = ckValue,
                Expires = DateTime.Now.AddHours(expireHours)
            };

            //Update
            HttpContext.Current.Response.Cookies.Set(responseCookie);
        }
        #endregion

        #region -- 其他功能 --
        /// <summary>
        /// 開始發信
        /// </summary>
        /// <param name="sender">寄件人</param>
        /// <param name="senderName">寄件人名稱</param>
        /// <param name="mailList">收件人清單</param>
        /// <param name="subject">主旨</param>
        /// <param name="mailBody">BODY</param>
        /// <param name="ErrMsg"></param>
        /// <returns></returns>
        public static bool Send_Email(string sender, string senderName, ArrayList mailList, string subject
            , StringBuilder mailBody, out string ErrMsg)
        {
            try
            {
                //開始發信
                using (MailMessage Msg = new MailMessage())
                {
                    //寄件人
                    Msg.From = new MailAddress(sender, senderName);

                    //收件人
                    foreach (string email in mailList)
                    {
                        Msg.To.Add(new MailAddress(email));
                    }

                    //主旨
                    Msg.Subject = subject;

                    //Body:取得郵件內容
                    Msg.Body = mailBody.ToString();

                    Msg.IsBodyHtml = true;

                    SmtpClient smtp = new SmtpClient();

                    smtp.Send(Msg);
                    smtp.Dispose();

                    //OK
                    ErrMsg = "";
                    return true;
                }
            }
            catch (Exception ex)
            {
                ErrMsg = "郵件發送失敗..." + ex.Message.ToString();
                return false;
            }
        }
        #endregion

        #region -- 字串過濾 --
        /// <summary>
        /// 過濾標記
        /// </summary>
        /// <param name="Htmlstring">包括HTML，指令碼，資料庫關鍵字，特殊字元的原始碼 </param>
        /// <returns>已經去除標記後的文字</returns>
        public static string filterHtml(string Htmlstring)
        {
            if (Htmlstring == null)
            {
                return "";
            }
            else
            {
                //刪除指令碼
                Htmlstring = Regex.Replace(Htmlstring, @"<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase);
                //刪除HTML
                Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"-->", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"<!--.*", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", " ", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, @"&#(\d+);", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "xp_cmdshell", "", RegexOptions.IgnoreCase);

                //刪除與資料庫相關的詞
                Htmlstring = Regex.Replace(Htmlstring, "select", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "insert", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "delete from", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "count''", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "drop table", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "truncate", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "asc", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "mid", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "char", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "xp_cmdshell", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "exec master", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "net localgroup administrators", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "and", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "net user", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "or", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "net", "", RegexOptions.IgnoreCase);
                //Htmlstring = Regex.Replace(Htmlstring, "*", "", RegexOptions.IgnoreCase);
                //Htmlstring = Regex.Replace(Htmlstring, "-", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "delete", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "drop", "", RegexOptions.IgnoreCase);
                Htmlstring = Regex.Replace(Htmlstring, "script", "", RegexOptions.IgnoreCase);

                //特殊的字元
                Htmlstring = Htmlstring.Replace("<", "");
                Htmlstring = Htmlstring.Replace(">", "");
                Htmlstring = Htmlstring.Replace("*", "");
                //Htmlstring = Htmlstring.Replace("-", "");
                Htmlstring = Htmlstring.Replace("?", "");
                Htmlstring = Htmlstring.Replace("'", "''");
                //Htmlstring = Htmlstring.Replace(",", "");
                //Htmlstring = Htmlstring.Replace("/", "");
                Htmlstring = Htmlstring.Replace(";", "");
                Htmlstring = Htmlstring.Replace("*/", "");
                Htmlstring = Htmlstring.Replace("\r\n", "");
                Htmlstring = HttpContext.Current.Server.HtmlEncode(Htmlstring).Trim();

                return Htmlstring;
            }
        }

        #endregion
    }
}
