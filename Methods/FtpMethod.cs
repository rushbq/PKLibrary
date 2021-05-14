using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace PKLib_Method.Methods
{
    public class FtpMethod : FtpBase
    {
        /// <summary>
        /// 設定FTP 連線字串
        /// </summary>
        /// <param name="userName">帳號</param>
        /// <param name="passWord">密碼</param>
        /// <param name="serverUrl">FTP路徑</param>
        public FtpMethod(string userName, string passWord, string serverUrl)
            : base(userName, passWord, serverUrl)
        {
        }

        #region -- FTP功能 --

        /// <summary>
        /// 上傳檔案到FTP
        /// </summary>
        /// <param name="myFile">HttpPostedFile</param>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <param name="fileName">檔名</param>
        /// <returns></returns>
        public bool FTP_doUpload(HttpPostedFile myFile, string uploadFolder, string fileName)
        {
            return FTP_doUpload(myFile, uploadFolder, fileName
                , false, 0, 0);
        }


        /// <summary>
        /// 上傳檔案到FTP
        /// </summary>
        /// <param name="myFile">HttpPostedFile</param>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <param name="fileName">檔名</param>
        /// <param name="resizeImg">是否重設圖片大小</param>
        /// <param name="resizeW">設定寬</param>
        /// <param name="resizeH">設定高</param>
        /// <returns></returns>
        public bool FTP_doUpload(HttpPostedFile myFile, string uploadFolder, string fileName
            , bool resizeImg, int resizeW, int resizeH)
        {
            try
            {
                //取得完整路徑
                string ftpUrl = string.Format("{0}{1}/{2}", this.ServerUrl, uploadFolder, fileName);

                //讀取上傳檔案,並轉成byte
                Stream streamObj = myFile.InputStream;

                //宣告byte
                Byte[] buffer;


                #region -- 檔案判別與處理 --

                //取得副檔名
                string GetExt = Path.GetExtension(myFile.FileName);

                //判斷是否為圖片
                switch (GetExt.ToLower())
                {
                    case ".jpg":
                    case ".png":
                    case ".jpeg":
                    case ".bmp":
                    case ".gif":
                        if (resizeImg)
                        {
                            //執行圖片壓縮
                            ImageMethod _img = new ImageMethod();
                            buffer = _img.reSizeImage(streamObj, resizeW, resizeH);
                        }
                        else
                        {
                            buffer = new Byte[myFile.ContentLength];
                        }

                        break;

                    default:
                        //其他檔案類型
                        buffer = new Byte[myFile.ContentLength];

                        break;
                }

                #endregion


                streamObj.Read(buffer, 0, buffer.Length);
                streamObj.Close();
                streamObj = null;

                //取得FTP協定
                FtpWebRequest requestObj = FtpWebRequest.Create(ftpUrl) as FtpWebRequest;

                //完成後,連線關閉
                requestObj.KeepAlive = false;
                requestObj.UseBinary = true;

                //method = 上傳
                requestObj.Method = WebRequestMethods.Ftp.UploadFile;
                requestObj.Credentials = new NetworkCredential(this.Username, this.Password);

                //上傳資料流
                Stream requestStream = requestObj.GetRequestStream();
                requestStream.Write(buffer, 0, buffer.Length);
                requestStream.Flush();
                requestStream.Close();
                requestObj = null;

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 上傳檔案到FTP
        /// 複製原來的檔案轉為資料流, 同張圖上傳要產生縮圖時使用
        /// </summary>
        /// <param name="myFile">HttpPostedFile</param>
        /// <param name="streamObj">複製的資料流</param>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <param name="fileName">檔名</param>
        /// <param name="resizeImg">是否重設圖片大小</param>
        /// <param name="resizeW">設定寬</param>
        /// <param name="resizeH">設定高</param>
        /// <returns></returns>
        /// <remarks>
        /// 使用範例:PKUpload/myProd/PicUpload.aspx
        /// </remarks>
        public bool FTP_doUpload(HttpPostedFile myFile, Stream streamObj, string uploadFolder, string fileName
           , bool resizeImg, int resizeW, int resizeH)
        {
            try
            {
                //取得完整路徑
                string ftpUrl = string.Format("{0}{1}/{2}", this.ServerUrl, uploadFolder, fileName);

                //讀取上傳檔案,並轉成byte
                //Stream streamObj = myFile.InputStream;

                //宣告byte
                Byte[] buffer;


                #region -- 檔案判別與處理 --
                
                //取得副檔名
                string GetExt = Path.GetExtension(myFile.FileName);

                //判斷是否為圖片
                switch (GetExt.ToLower())
                {
                    case ".jpg":
                    case ".png":
                    case ".jpeg":
                    case ".bmp":
                    case ".gif":
                        if (resizeImg)
                        {
                            //執行圖片壓縮
                            ImageMethod _img = new ImageMethod();
                            buffer = _img.reSizeImage(streamObj, resizeW, resizeH);
                        }
                        else
                        {
                            buffer = new Byte[myFile.ContentLength];
                        }

                        break;

                    default:
                        //其他檔案類型
                        buffer = new Byte[myFile.ContentLength];

                        break;
                }

                #endregion


                streamObj.Read(buffer, 0, buffer.Length);
                streamObj.Close();
                streamObj = null;

                //取得FTP協定
                FtpWebRequest requestObj = FtpWebRequest.Create(ftpUrl) as FtpWebRequest;

                //完成後,連線關閉
                requestObj.KeepAlive = false;
                requestObj.UseBinary = true;

                //method = 上傳
                requestObj.Method = WebRequestMethods.Ftp.UploadFile;
                requestObj.Credentials = new NetworkCredential(this.Username, this.Password);

                //上傳資料流
                Stream requestStream = requestObj.GetRequestStream();
                requestStream.Write(buffer, 0, buffer.Length);
                requestStream.Flush();
                requestStream.Close();
                requestObj = null;

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        /// <summary>
        /// 上傳檔案到FTP, 使用轉好的byte
        /// </summary>
        /// <param name="myfile">byte</param>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <param name="fileName">檔名</param>
        /// <returns></returns>
        public bool FTP_doUploadWithByte(byte[] myfile, string uploadFolder, string fileName)
        {
            try
            {
                //取得完整路徑
                string ftpUrl = string.Format("{0}{1}/{2}", this.ServerUrl, uploadFolder, fileName);
                
                //取得FTP協定
                FtpWebRequest requestObj = FtpWebRequest.Create(ftpUrl) as FtpWebRequest;

                //完成後,連線關閉
                requestObj.KeepAlive = false;
                requestObj.UseBinary = true;

                //method = 上傳
                requestObj.Method = WebRequestMethods.Ftp.UploadFile;
                requestObj.Credentials = new NetworkCredential(this.Username, this.Password);

                //上傳資料流
                Stream requestStream = requestObj.GetRequestStream();
                requestStream.Write(myfile, 0, myfile.Length);
                requestStream.Flush();
                requestStream.Close();
                requestObj = null;

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// 從FTP下載檔案
        /// </summary>
        /// <param name="uploadFolder">FTP資料夾</param>
        /// <param name="realFileName">原始檔名</param>
        /// <param name="dwFileName">另存新檔的檔名</param>
        /// <example>
        /// 直接在Postback事件裡使用
        /// </example>
        public void FTP_doDownload(string uploadFolder, string realFileName, string dwFileName)
        {
            try
            {
                //取得完整路徑
                string ftpUrl = string.Format("{0}{1}/{2}", this.ServerUrl, uploadFolder, realFileName);

                //取得FTP協定
                FtpWebRequest dwRequest = (FtpWebRequest)WebRequest.Create(ftpUrl);

                //完成後,連線關閉
                dwRequest.KeepAlive = false;
                dwRequest.UseBinary = true;

                //method = 下載
                dwRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                dwRequest.Credentials = new NetworkCredential(this.Username, this.Password);

                //dwRequest.Timeout = (60000 * 1);  // (60000 * 1) 一分鐘

                //取得FTP伺服器回應
                using (FtpWebResponse dwResponse = (FtpWebResponse)dwRequest.GetResponse())
                {
                    //取得FTP伺服器回傳的資料流
                    using (Stream responseStream = dwResponse.GetResponseStream())
                    {
                        //取得資料流
                        using (StreamReader reader = new StreamReader(responseStream))
                        {
                            string fileName = Path.GetFileName(dwRequest.RequestUri.AbsolutePath);
                            if (fileName.Length == 0)
                            {
                                throw new Exception(reader.ReadToEnd());
                            }
                            else
                            {
                                int Length = 2048;
                                byte[] buffer = new byte[Length];
                                int read = 0;

                                using (MemoryStream ms = new MemoryStream())
                                {
                                    while ((read = responseStream.Read(buffer, 0, Length)) > 0)
                                    {
                                        ms.Write(buffer, 0, read);
                                    }
                                    ms.Flush();
                                    HttpContext.Current.Response.Clear();
                                    HttpContext.Current.Response.ContentType = "application/octet-stream";

                                    // 設定強制下載標頭
                                    HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=" + dwFileName));
                                    // 輸出檔案
                                    HttpContext.Current.Response.BinaryWrite(ms.ToArray());
                                }
                            }
                        }

                    }

                }

                dwRequest = null;

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 判斷資料夾是否存在
        /// </summary>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <returns></returns>
        public void FTP_CheckFolder(string uploadFolder)
        {
            try
            {
                //宣告
                FtpWebRequest reqFTP = null;
                Stream ftpStream = null;

                //目前根目錄
                string currentDir = this.ServerUrl;

                //解析來源資料夾, 解決多層目錄問題
                string[] subDirs = uploadFolder.Split('/');

                //建立資料夾
                foreach (string subDir in subDirs)
                {
                    try
                    {
                        currentDir = currentDir + subDir + "/";
                        reqFTP = (FtpWebRequest)FtpWebRequest.Create(currentDir);
                        reqFTP.Method = WebRequestMethods.Ftp.MakeDirectory;
                        reqFTP.UseBinary = true;
                        //登入
                        reqFTP.Credentials = new NetworkCredential(this.Username, this.Password);
                        FtpWebResponse response = (FtpWebResponse)reqFTP.GetResponse();
                        ftpStream = response.GetResponseStream();
                        ftpStream.Close();
                        response.Close();
                    }
                    catch (Exception ex)
                    {
                        //directory already exist I know that is weak but there is no way to check if a folder exist on ftp...
                    }
                }

            }
            catch (Exception)
            {
                //資料夾已存在,不執行任何動作
            }

        }

        /// <summary>
        /// 判斷檔案是否存在
        /// </summary>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <param name="fileName">檔案名稱</param>
        /// <returns></returns>
        public bool FTP_CheckFile(string uploadFolder, string fileName)
        {
            try
            {
                string myUrl = this.ServerUrl + uploadFolder + @"/" + fileName;

                //宣告
                FtpWebRequest ftp = (FtpWebRequest)FtpWebRequest.Create(myUrl);

                //登入
                ftp.Credentials = new NetworkCredential(this.Username, this.Password);

                //取得檔案大小
                ftp.Method = WebRequestMethods.Ftp.GetFileSize;

                //取得FTP回應
                FtpWebResponse myResponse = (FtpWebResponse)ftp.GetResponse();

                string result = string.Empty;

                using (Stream datastream = myResponse.GetResponseStream())
                {
                    StreamReader sr = new StreamReader(datastream);
                    result = sr.ReadToEnd();
                    sr.Close();
                }

                myResponse.Close();
                ftp = null;

                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        /// <summary>
        /// 列出資料夾內的檔案
        /// </summary>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <returns></returns>
        public List<string> ListFiles(string uploadFolder)
        {
            string myUrl = this.ServerUrl + uploadFolder;

            //宣告
            FtpWebRequest ftp = (FtpWebRequest)FtpWebRequest.Create(myUrl);

            //登入
            ftp.Credentials = new NetworkCredential(this.Username, this.Password);

            //顯示資料夾內容檔案
            ftp.Method = WebRequestMethods.Ftp.ListDirectory;

            //取得FTP回應
            FtpWebResponse myResponse = (FtpWebResponse)ftp.GetResponse();

            List<string> result = new List<string>();

            using (Stream datastream = myResponse.GetResponseStream())
            {
                StreamReader reader = new StreamReader(datastream);

                while (!reader.EndOfStream)
                {
                    result.Add(reader.ReadLine());
                }

                reader.Close();
            }

            ftp = null;

            return result;
        }

        /// <summary>
        /// 刪除檔案
        /// </summary>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <param name="fileName">檔案名稱</param>
        /// <returns></returns>
        public bool FTP_DelFile(string uploadFolder, string fileName)
        {
            try
            {
                string myUrl = this.ServerUrl + uploadFolder + @"/" + fileName;

                //宣告
                FtpWebRequest ftp = (FtpWebRequest)FtpWebRequest.Create(myUrl);

                //登入
                ftp.Credentials = new NetworkCredential(this.Username, this.Password);

                //刪除檔案
                ftp.Method = WebRequestMethods.Ftp.DeleteFile;

                //取得FTP回應
                FtpWebResponse myResponse = (FtpWebResponse)ftp.GetResponse();
                string result = string.Empty;

                using (Stream datastream = myResponse.GetResponseStream())
                {
                    StreamReader sr = new StreamReader(datastream);
                    result = sr.ReadToEnd();
                    sr.Close();
                }

                myResponse.Close();
                ftp = null;

                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        /// <summary>
        /// 刪除資料夾
        /// </summary>
        /// <param name="uploadFolder">資料夾名稱</param>
        /// <returns></returns>
        public bool FTP_DelFolder(string uploadFolder)
        {
            try
            {
                string myUrl = this.ServerUrl + uploadFolder;

                //宣告
                FtpWebRequest ftp = (FtpWebRequest)FtpWebRequest.Create(myUrl);

                //登入
                ftp.Credentials = new NetworkCredential(this.Username, this.Password);

                //取得資料夾內的檔案, 並刪除所有檔案
                List<string> filesList = ListFiles(uploadFolder);
                foreach (string file in filesList)
                {
                    FTP_DelFile(uploadFolder, file);
                }

                //刪除資料夾
                ftp.Method = WebRequestMethods.Ftp.RemoveDirectory;

                //取得FTP回應
                FtpWebResponse myResponse = (FtpWebResponse)ftp.GetResponse();

                myResponse.Close();
                ftp = null;

                return true;

            }
            catch (Exception)
            {
                return false;
            }

        }


        /// <summary>
        /// 旋轉圖片
        /// </summary>
        /// <param name="uploadFolder">FTP目錄</param>
        /// <param name="realFileName">實際檔名</param>
        /// <param name="degree">旋轉角度</param>
        public bool RotateImage(string uploadFolder, string realFileName, float degree)
        {
            try
            {
                //取得完整路徑
                string ftpUrl = string.Format("{0}{1}/{2}", this.ServerUrl, uploadFolder, realFileName);

                //宣告byte(上傳使用)
                Byte[] ImgBuffer;

                #region -- 取得檔案, 旋轉圖片 --

                //取得FTP協定
                FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create(ftpUrl);

                //完成後,連線關閉
                ftpRequest.KeepAlive = false;
                ftpRequest.UseBinary = true;

                //[FTP] - 下載
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                ftpRequest.Credentials = new NetworkCredential(this.Username, this.Password);


                //取得FTP伺服器回應
                using (FtpWebResponse dwResponse = (FtpWebResponse)ftpRequest.GetResponse())
                {
                    //取得FTP伺服器回傳的資料流
                    using (Stream responseStream = dwResponse.GetResponseStream())
                    {
                        //執行圖片旋轉
                        ImageMethod _img = new ImageMethod();
                        ImgBuffer = _img.rotateImage(responseStream, degree);
                    }
                }

                ftpRequest = null;

                #endregion


                #region -- 重新上傳照片 --

                //[FTP] - 上傳
                //取得FTP協定
                FtpWebRequest requestObj = FtpWebRequest.Create(ftpUrl) as FtpWebRequest;

                //完成後,連線關閉
                requestObj.KeepAlive = false;
                requestObj.UseBinary = true;

                //method = 上傳
                requestObj.Method = WebRequestMethods.Ftp.UploadFile;
                requestObj.Credentials = new NetworkCredential(this.Username, this.Password);

                //上傳資料流
                Stream requestStream = requestObj.GetRequestStream();
                requestStream.Write(ImgBuffer, 0, ImgBuffer.Length);
                requestStream.Flush();
                requestStream.Close();
                requestObj = null;

                #endregion

                return true;

            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion


    }
}
