using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace cardsales
{
    class Program
    {
        static RfcDestination m_rfcDestination = null;
        static RfcRepository m_rfcRepository = null;
        static IRfcFunction m_rfcFunction;
        static DataTable posZC01 = new DataTable();  // 01 포스 + ZC01 비씨카드
        static DataTable posZC02 = new DataTable();  // 01 포스 + ZC02 하나카드
        static DataTable posZC03 = new DataTable();  // 01 포스 + ZC03 신한카드
        static DataTable posZC04 = new DataTable();  // 01 포스 + ZC04 국민카드
        static DataTable posZC05 = new DataTable();  // 01 포스 + ZC05 삼성카드
        static DataTable posZC07 = new DataTable();  // 01 포스 + ZC07 현대카드
        static DataTable posZC08 = new DataTable();  // 01 포스 + ZC08 롯데카드
        static DataTable easyZC01 = new DataTable(); // 02 이지 + ZC01 비씨카드
        static DataTable easyZC02 = new DataTable(); // 02 이지 + ZC02 하나카드
        static DataTable easyZC03 = new DataTable(); // 02 이지 + ZC03 신한카드
        static DataTable easyZC04 = new DataTable(); // 02 이지 + ZC04 국민카드
        static DataTable easyZC05 = new DataTable(); // 02 이지 + ZC05 삼성카드
        static DataTable easyZC07 = new DataTable(); // 02 이지 + ZC07 현대카드
        static DataTable easyZC08 = new DataTable(); // 02 이지 + ZC08 롯데카드
        static DataTable sangZC01 = new DataTable(); // 03 상품권 + ZC01 비씨카드
        static DataTable sangZC02 = new DataTable(); // 03 상품권 + ZC02 하나카드
        static DataTable sangZC03 = new DataTable(); // 03 상품권 + ZC03 신한카드
        static DataTable sangZC04 = new DataTable(); // 03 상품권 + ZC04 국민카드
        static DataTable sangZC05 = new DataTable(); // 03 상품권 + ZC05 삼성카드
        static DataTable sangZC07 = new DataTable(); // 03 상품권 + ZC07 현대카드
        static DataTable sangZC08 = new DataTable(); // 03 상품권 + ZC08 롯데카드
        static string targetDate = string.Empty;
        static string targetDateStart = string.Empty;
        static string targetDateEnd = string.Empty;

        // Selenium 세팅
        static ChromeDriverService service = null;
        static ChromeOptions option = null;
        static ChromeDriver driver = null;

        static void Main(string[] args)
        {
            if (args.Length == 0) // 자동 (날짜 인수 없을 때)
            {
                targetDate = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");
                ParsingData(targetDate);
                RFCInfoSet("여신협회 가맹점 입금정보 인터페이스", "*");

                // 01 포스 + ZC01 비씨카드
                InsertRfcDataImport("I_GUBUN", "01", "I_ZCCCD", "ZC01");
                string l_strTableName = "T_DOCUMENT";
                InsertRfcDataRows("*", l_strTableName, posZC01);
                RfcFunctionCall();
                string l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                string l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 01 포스 + ZC02 하나카드
                InsertRfcDataImport("I_GUBUN", "01", "I_ZCCCD", "ZC02");
                InsertRfcDataRows("*", l_strTableName, posZC02);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 01 포스 + ZC03 신한카드
                InsertRfcDataImport("I_GUBUN", "01", "I_ZCCCD", "ZC03");
                InsertRfcDataRows("*", l_strTableName, posZC03);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 01 포스 + ZC04 국민카드
                InsertRfcDataImport("I_GUBUN", "01", "I_ZCCCD", "ZC04");
                InsertRfcDataRows("*", l_strTableName, posZC04);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 01 포스 + ZC05 삼성카드
                InsertRfcDataImport("I_GUBUN", "01", "I_ZCCCD", "ZC05");
                InsertRfcDataRows("*", l_strTableName, posZC05);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 01 포스 + ZC07 현대카드
                InsertRfcDataImport("I_GUBUN", "01", "I_ZCCCD", "ZC07");
                InsertRfcDataRows("*", l_strTableName, posZC07);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 01 포스 + ZC08 롯데카드
                InsertRfcDataImport("I_GUBUN", "01", "I_ZCCCD", "ZC08");
                InsertRfcDataRows("*", l_strTableName, posZC08);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);

                // 02 이지 + ZC01 비씨카드
                InsertRfcDataImport("I_GUBUN", "02", "I_ZCCCD", "ZC01");
                InsertRfcDataRows("*", l_strTableName, easyZC01);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 02 이지 + ZC02 하나카드
                InsertRfcDataImport("I_GUBUN", "02", "I_ZCCCD", "ZC02");
                InsertRfcDataRows("*", l_strTableName, easyZC02);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 02 이지 + ZC03 신한카드
                InsertRfcDataImport("I_GUBUN", "02", "I_ZCCCD", "ZC03");
                InsertRfcDataRows("*", l_strTableName, easyZC03);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 02 이지 + ZC04 국민카드
                InsertRfcDataImport("I_GUBUN", "02", "I_ZCCCD", "ZC04");
                InsertRfcDataRows("*", l_strTableName, easyZC04);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 02 이지 + ZC05 삼성카드
                InsertRfcDataImport("I_GUBUN", "02", "I_ZCCCD", "ZC05");
                InsertRfcDataRows("*", l_strTableName, easyZC05);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 02 이지 + ZC07 현대카드
                InsertRfcDataImport("I_GUBUN", "02", "I_ZCCCD", "ZC07");
                InsertRfcDataRows("*", l_strTableName, easyZC07);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 02 이지 + ZC08 롯데카드
                InsertRfcDataImport("I_GUBUN", "02", "I_ZCCCD", "ZC08");
                InsertRfcDataRows("*", l_strTableName, easyZC08);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);

                // 03 상품권 + ZC01 비씨카드
                InsertRfcDataImport("I_GUBUN", "03", "I_ZCCCD", "ZC01");
                InsertRfcDataRows("*", l_strTableName, sangZC01);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 03 상품권 + ZC02 하나카드
                InsertRfcDataImport("I_GUBUN", "03", "I_ZCCCD", "ZC02");
                InsertRfcDataRows("*", l_strTableName, sangZC02);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 03 상품권 + ZC03 신한카드
                InsertRfcDataImport("I_GUBUN", "03", "I_ZCCCD", "ZC03");
                InsertRfcDataRows("*", l_strTableName, sangZC03);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 03 상품권 + ZC04 국민카드
                InsertRfcDataImport("I_GUBUN", "03", "I_ZCCCD", "ZC04");
                InsertRfcDataRows("*", l_strTableName, sangZC04);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 03 상품권 + ZC05 삼성카드
                InsertRfcDataImport("I_GUBUN", "03", "I_ZCCCD", "ZC05");
                InsertRfcDataRows("*", l_strTableName, sangZC05);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 03 상품권 + ZC07 현대카드
                InsertRfcDataImport("I_GUBUN", "03", "I_ZCCCD", "ZC07");
                InsertRfcDataRows("*", l_strTableName, sangZC07);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);
                // 03 상품권 + ZC08 롯데카드
                InsertRfcDataImport("I_GUBUN", "03", "I_ZCCCD", "ZC08");
                InsertRfcDataRows("*", l_strTableName, sangZC08);
                RfcFunctionCall();
                l_strErrorCode = m_rfcFunction.GetValue("E_RESULT").ToString();
                l_strErrorMessage = m_rfcFunction.GetValue("E_MESSAGE").ToString();
                WriteLog("RFC 결과 : " + l_strErrorCode + "\t" + l_strErrorMessage);

                WriteLog("프로그램 종료");
            }
            //else if (args.Length == 1) // 수동 (날짜 인수 1개 있을 때 : 특정 날짜만)
            //{
            //    _targetDate = args[0].ToString();
            //    ParsingData(_targetDate);
            //}
            //else if (args.Length == 2) // 수동 (날짜 인수 2개 있을 때 : 특정 기간동안)
            //{
            //    _targetDateStart = args[0].ToString();
            //    _targetDateEnd = args[1].ToString();
            //    _targetDate = _targetDateStart;
            //}
        }

        private static void ParsingData(string _targetDate)
        {
            // DataTable 세팅
            #region
            posZC01.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            posZC01.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            posZC01.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            posZC01.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            posZC02.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            posZC02.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            posZC02.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            posZC02.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            posZC03.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            posZC03.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            posZC03.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            posZC03.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            posZC04.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            posZC04.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            posZC04.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            posZC04.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            posZC05.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            posZC05.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            posZC05.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            posZC05.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            posZC07.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            posZC07.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            posZC07.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            posZC07.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            posZC08.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            posZC08.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            posZC08.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            posZC08.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            easyZC01.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            easyZC01.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            easyZC01.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            easyZC01.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            easyZC02.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            easyZC02.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            easyZC02.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            easyZC02.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            easyZC03.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            easyZC03.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            easyZC03.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            easyZC03.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            easyZC04.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            easyZC04.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            easyZC04.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            easyZC04.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            easyZC05.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            easyZC05.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            easyZC05.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            easyZC05.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            easyZC07.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            easyZC07.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            easyZC07.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            easyZC07.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            easyZC08.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            easyZC08.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            easyZC08.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            easyZC08.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            sangZC01.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            sangZC01.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            sangZC01.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            sangZC01.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            sangZC02.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            sangZC02.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            sangZC02.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            sangZC02.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            sangZC03.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            sangZC03.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            sangZC03.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            sangZC03.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            sangZC04.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            sangZC04.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            sangZC04.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            sangZC04.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            sangZC05.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            sangZC05.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            sangZC05.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            sangZC05.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            sangZC07.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            sangZC07.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            sangZC07.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            sangZC07.Columns.Add(new DataColumn("BKINPUT", typeof(string)));
            sangZC08.Columns.Add(new DataColumn("BKDATE", typeof(string)));
            sangZC08.Columns.Add(new DataColumn("SALCNT", typeof(string)));
            sangZC08.Columns.Add(new DataColumn("SALAMT", typeof(string)));
            sangZC08.Columns.Add(new DataColumn("BKINPUT", typeof(string)));

            string[] dr = new string[4];
            #endregion

            try // ChromeDriver 오류 방지용 try 구문
            {
                // Selenium 세팅
                service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
                option = new ChromeOptions();
                option.AddArguments("disable-gpu", "headless", "no-sandbox");
                driver = new ChromeDriver(service, option, TimeSpan.FromMinutes(3));
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(3);
            }
            catch (Exception ex)
            {
                WriteLog("*ChromeDriver 오류* " + ex.Message);
                driver.Quit(); // ChromeDriver 종료
            }

            try // 인터넷 오류 방지용 try 구문
            {
                // 여신금융협회 로그인 -> 입금내역 조회 페이지로 이동
                driver.Url = "https://www.cardsales.or.kr/signin";
            }
            catch (Exception ex)
            {
                WriteLog("*여신금융협회 접속 오류* " + ex.Message);
                driver.Quit(); // ChromeDriver 종료
            }

            var element = driver.FindElementById("j_username");
            element.SendKeys("*");
            element = driver.FindElementById("j_password");
            element.SendKeys("*");
            element = driver.FindElementById("goLogin");
            element.Click();
            element = driver.FindElementByXPath("//*[@id='wrap']/header/div[2]/nav/ul/li[3]/a"); // 상단 입금내역 버튼
            element.Click();

            // <--- 입금내역 조회 페이지 --->
            #region 01 포스 + ZC01 비씨카드 (4)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[1]"); // 가맹점그룹 - 영풍문고(pos)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[4]"); // 카드사 - 비씨카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(pos) + 비씨카드 *****");
                WriteLog("***** 영풍문고(pos) + 비씨카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                posZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC01.Rows[0][0] + "\t" + "매출건수: " + posZC01.Rows[0][1] + "\t" + "매출합계: " + posZC01.Rows[0][2] + "\t" + "입금합계: " + posZC01.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC01.Rows[0][0] + "\t" + "매출건수: " + posZC01.Rows[0][1] + "\t" + "매출합계: " + posZC01.Rows[0][2] + "\t" + "입금합계: " + posZC01.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                posZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC01.Rows[1][0] + "\t" + "매출건수: " + posZC01.Rows[1][1] + "\t" + "매출합계: " + posZC01.Rows[1][2] + "\t" + "입금합계: " + posZC01.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC01.Rows[1][0] + "\t" + "매출건수: " + posZC01.Rows[1][1] + "\t" + "매출합계: " + posZC01.Rows[1][2] + "\t" + "입금합계: " + posZC01.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                posZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC01.Rows[2][0] + "\t" + "매출건수: " + posZC01.Rows[2][1] + "\t" + "매출합계: " + posZC01.Rows[2][2] + "\t" + "입금합계: " + posZC01.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC01.Rows[2][0] + "\t" + "매출건수: " + posZC01.Rows[2][1] + "\t" + "매출합계: " + posZC01.Rows[2][2] + "\t" + "입금합계: " + posZC01.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 01 포스 + ZC02 하나카드 (10)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[1]"); // 가맹점그룹 - 영풍문고(pos)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[10]"); // 카드사 - 하나카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(pos) + 하나카드 *****");
                WriteLog("***** 영풍문고(pos) + 하나카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                posZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC02.Rows[0][0] + "\t" + "매출건수: " + posZC02.Rows[0][1] + "\t" + "매출합계: " + posZC02.Rows[0][2] + "\t" + "입금합계: " + posZC02.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC02.Rows[0][0] + "\t" + "매출건수: " + posZC02.Rows[0][1] + "\t" + "매출합계: " + posZC02.Rows[0][2] + "\t" + "입금합계: " + posZC02.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                posZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC02.Rows[1][0] + "\t" + "매출건수: " + posZC02.Rows[1][1] + "\t" + "매출합계: " + posZC02.Rows[1][2] + "\t" + "입금합계: " + posZC02.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC02.Rows[1][0] + "\t" + "매출건수: " + posZC02.Rows[1][1] + "\t" + "매출합계: " + posZC02.Rows[1][2] + "\t" + "입금합계: " + posZC02.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                posZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC02.Rows[2][0] + "\t" + "매출건수: " + posZC02.Rows[2][1] + "\t" + "매출합계: " + posZC02.Rows[2][2] + "\t" + "입금합계: " + posZC02.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC02.Rows[2][0] + "\t" + "매출건수: " + posZC02.Rows[2][1] + "\t" + "매출합계: " + posZC02.Rows[2][2] + "\t" + "입금합계: " + posZC02.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 01 포스 + ZC03 신한카드 (3)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[1]"); // 가맹점그룹 - 영풍문고(pos)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[3]"); // 카드사 - 신한카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(pos) + 신한카드 *****");
                WriteLog("***** 영풍문고(pos) + 신한카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                posZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC03.Rows[0][0] + "\t" + "매출건수: " + posZC03.Rows[0][1] + "\t" + "매출합계: " + posZC03.Rows[0][2] + "\t" + "입금합계: " + posZC03.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC03.Rows[0][0] + "\t" + "매출건수: " + posZC03.Rows[0][1] + "\t" + "매출합계: " + posZC03.Rows[0][2] + "\t" + "입금합계: " + posZC03.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                posZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC03.Rows[1][0] + "\t" + "매출건수: " + posZC03.Rows[1][1] + "\t" + "매출합계: " + posZC03.Rows[1][2] + "\t" + "입금합계: " + posZC03.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC03.Rows[1][0] + "\t" + "매출건수: " + posZC03.Rows[1][1] + "\t" + "매출합계: " + posZC03.Rows[1][2] + "\t" + "입금합계: " + posZC03.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                posZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC03.Rows[2][0] + "\t" + "매출건수: " + posZC03.Rows[2][1] + "\t" + "매출합계: " + posZC03.Rows[2][2] + "\t" + "입금합계: " + posZC03.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC03.Rows[2][0] + "\t" + "매출건수: " + posZC03.Rows[2][1] + "\t" + "매출합계: " + posZC03.Rows[2][2] + "\t" + "입금합계: " + posZC03.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 01 포스 + ZC04 국민카드 (2)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[1]"); // 가맹점그룹 - 영풍문고(pos)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[2]"); // 카드사 - KB카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(pos) + 국민카드 *****");
                WriteLog("***** 영풍문고(pos) + 국민카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                posZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC04.Rows[0][0] + "\t" + "매출건수: " + posZC04.Rows[0][1] + "\t" + "매출합계: " + posZC04.Rows[0][2] + "\t" + "입금합계: " + posZC04.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC04.Rows[0][0] + "\t" + "매출건수: " + posZC04.Rows[0][1] + "\t" + "매출합계: " + posZC04.Rows[0][2] + "\t" + "입금합계: " + posZC04.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                posZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC04.Rows[1][0] + "\t" + "매출건수: " + posZC04.Rows[1][1] + "\t" + "매출합계: " + posZC04.Rows[1][2] + "\t" + "입금합계: " + posZC04.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC04.Rows[1][0] + "\t" + "매출건수: " + posZC04.Rows[1][1] + "\t" + "매출합계: " + posZC04.Rows[1][2] + "\t" + "입금합계: " + posZC04.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                posZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC04.Rows[2][0] + "\t" + "매출건수: " + posZC04.Rows[2][1] + "\t" + "매출합계: " + posZC04.Rows[2][2] + "\t" + "입금합계: " + posZC04.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC04.Rows[2][0] + "\t" + "매출건수: " + posZC04.Rows[2][1] + "\t" + "매출합계: " + posZC04.Rows[2][2] + "\t" + "입금합계: " + posZC04.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 01 포스 + ZC05 삼성카드 (7)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[1]"); // 가맹점그룹 - 영풍문고(pos)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[7]"); // 카드사 - 삼성카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(pos) + 삼성카드 *****");
                WriteLog("***** 영풍문고(pos) + 삼성카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                posZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC05.Rows[0][0] + "\t" + "매출건수: " + posZC05.Rows[0][1] + "\t" + "매출합계: " + posZC05.Rows[0][2] + "\t" + "입금합계: " + posZC05.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC05.Rows[0][0] + "\t" + "매출건수: " + posZC05.Rows[0][1] + "\t" + "매출합계: " + posZC05.Rows[0][2] + "\t" + "입금합계: " + posZC05.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                posZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC05.Rows[1][0] + "\t" + "매출건수: " + posZC05.Rows[1][1] + "\t" + "매출합계: " + posZC05.Rows[1][2] + "\t" + "입금합계: " + posZC05.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC05.Rows[1][0] + "\t" + "매출건수: " + posZC05.Rows[1][1] + "\t" + "매출합계: " + posZC05.Rows[1][2] + "\t" + "입금합계: " + posZC05.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                posZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC05.Rows[2][0] + "\t" + "매출건수: " + posZC05.Rows[2][1] + "\t" + "매출합계: " + posZC05.Rows[2][2] + "\t" + "입금합계: " + posZC05.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC05.Rows[2][0] + "\t" + "매출건수: " + posZC05.Rows[2][1] + "\t" + "매출합계: " + posZC05.Rows[2][2] + "\t" + "입금합계: " + posZC05.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 01 포스 + ZC07 현대카드 (6)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[1]"); // 가맹점그룹 - 영풍문고(pos)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[6]"); // 카드사 - 현대카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(pos) + 현대카드 *****");
                WriteLog("***** 영풍문고(pos) + 현대카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                posZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC07.Rows[0][0] + "\t" + "매출건수: " + posZC07.Rows[0][1] + "\t" + "매출합계: " + posZC07.Rows[0][2] + "\t" + "입금합계: " + posZC07.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC07.Rows[0][0] + "\t" + "매출건수: " + posZC07.Rows[0][1] + "\t" + "매출합계: " + posZC07.Rows[0][2] + "\t" + "입금합계: " + posZC07.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                posZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC07.Rows[1][0] + "\t" + "매출건수: " + posZC07.Rows[1][1] + "\t" + "매출합계: " + posZC07.Rows[1][2] + "\t" + "입금합계: " + posZC07.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC07.Rows[1][0] + "\t" + "매출건수: " + posZC07.Rows[1][1] + "\t" + "매출합계: " + posZC07.Rows[1][2] + "\t" + "입금합계: " + posZC07.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                posZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC07.Rows[2][0] + "\t" + "매출건수: " + posZC07.Rows[2][1] + "\t" + "매출합계: " + posZC07.Rows[2][2] + "\t" + "입금합계: " + posZC07.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC07.Rows[2][0] + "\t" + "매출건수: " + posZC07.Rows[2][1] + "\t" + "매출합계: " + posZC07.Rows[2][2] + "\t" + "입금합계: " + posZC07.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 01 포스 + ZC08 롯데카드 (5)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[1]"); // 가맹점그룹 - 영풍문고(pos)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[5]"); // 카드사 - 롯데카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(pos) + 롯데카드 *****");
                WriteLog("***** 영풍문고(pos) + 롯데카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                posZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC08.Rows[0][0] + "\t" + "매출건수: " + posZC08.Rows[0][1] + "\t" + "매출합계: " + posZC08.Rows[0][2] + "\t" + "입금합계: " + posZC08.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC08.Rows[0][0] + "\t" + "매출건수: " + posZC08.Rows[0][1] + "\t" + "매출합계: " + posZC08.Rows[0][2] + "\t" + "입금합계: " + posZC08.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                posZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC08.Rows[1][0] + "\t" + "매출건수: " + posZC08.Rows[1][1] + "\t" + "매출합계: " + posZC08.Rows[1][2] + "\t" + "입금합계: " + posZC08.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC08.Rows[1][0] + "\t" + "매출건수: " + posZC08.Rows[1][1] + "\t" + "매출합계: " + posZC08.Rows[1][2] + "\t" + "입금합계: " + posZC08.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                posZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + posZC08.Rows[2][0] + "\t" + "매출건수: " + posZC08.Rows[2][1] + "\t" + "매출합계: " + posZC08.Rows[2][2] + "\t" + "입금합계: " + posZC08.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + posZC08.Rows[2][0] + "\t" + "매출건수: " + posZC08.Rows[2][1] + "\t" + "매출합계: " + posZC08.Rows[2][2] + "\t" + "입금합계: " + posZC08.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion


            #region 02 이지 + ZC01 비씨카드 (4)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[2]"); // 가맹점그룹 - 영풍문고(이지)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[4]"); // 카드사 - 비씨카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(이지) + 비씨카드 *****");
                WriteLog("***** 영풍문고(이지) + 비씨카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC01.Rows[0][0] + "\t" + "매출건수: " + easyZC01.Rows[0][1] + "\t" + "매출합계: " + easyZC01.Rows[0][2] + "\t" + "입금합계: " + easyZC01.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC01.Rows[0][0] + "\t" + "매출건수: " + easyZC01.Rows[0][1] + "\t" + "매출합계: " + easyZC01.Rows[0][2] + "\t" + "입금합계: " + easyZC01.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC01.Rows[1][0] + "\t" + "매출건수: " + easyZC01.Rows[1][1] + "\t" + "매출합계: " + easyZC01.Rows[1][2] + "\t" + "입금합계: " + easyZC01.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC01.Rows[1][0] + "\t" + "매출건수: " + easyZC01.Rows[1][1] + "\t" + "매출합계: " + easyZC01.Rows[1][2] + "\t" + "입금합계: " + easyZC01.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC01.Rows[2][0] + "\t" + "매출건수: " + easyZC01.Rows[2][1] + "\t" + "매출합계: " + easyZC01.Rows[2][2] + "\t" + "입금합계: " + easyZC01.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC01.Rows[2][0] + "\t" + "매출건수: " + easyZC01.Rows[2][1] + "\t" + "매출합계: " + easyZC01.Rows[2][2] + "\t" + "입금합계: " + easyZC01.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 02 이지 + ZC02 하나카드 (10)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[2]"); // 가맹점그룹 - 영풍문고(이지)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[10]"); // 카드사 - 하나카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(이지) + 하나카드 *****");
                WriteLog("***** 영풍문고(이지) + 하나카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC02.Rows[0][0] + "\t" + "매출건수: " + easyZC02.Rows[0][1] + "\t" + "매출합계: " + easyZC02.Rows[0][2] + "\t" + "입금합계: " + easyZC02.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC02.Rows[0][0] + "\t" + "매출건수: " + easyZC02.Rows[0][1] + "\t" + "매출합계: " + easyZC02.Rows[0][2] + "\t" + "입금합계: " + easyZC02.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC02.Rows[1][0] + "\t" + "매출건수: " + easyZC02.Rows[1][1] + "\t" + "매출합계: " + easyZC02.Rows[1][2] + "\t" + "입금합계: " + easyZC02.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC02.Rows[1][0] + "\t" + "매출건수: " + easyZC02.Rows[1][1] + "\t" + "매출합계: " + easyZC02.Rows[1][2] + "\t" + "입금합계: " + easyZC02.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC02.Rows[2][0] + "\t" + "매출건수: " + easyZC02.Rows[2][1] + "\t" + "매출합계: " + easyZC02.Rows[2][2] + "\t" + "입금합계: " + easyZC02.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC02.Rows[2][0] + "\t" + "매출건수: " + easyZC02.Rows[2][1] + "\t" + "매출합계: " + easyZC02.Rows[2][2] + "\t" + "입금합계: " + easyZC02.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 02 이지 + ZC03 신한카드 (3)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[2]"); // 가맹점그룹 - 영풍문고(이지)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[3]"); // 카드사 - 신한카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(이지) + 신한카드 *****");
                WriteLog("***** 영풍문고(이지) + 신한카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC03.Rows[0][0] + "\t" + "매출건수: " + easyZC03.Rows[0][1] + "\t" + "매출합계: " + easyZC03.Rows[0][2] + "\t" + "입금합계: " + easyZC03.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC03.Rows[0][0] + "\t" + "매출건수: " + easyZC03.Rows[0][1] + "\t" + "매출합계: " + easyZC03.Rows[0][2] + "\t" + "입금합계: " + easyZC03.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC03.Rows[1][0] + "\t" + "매출건수: " + easyZC03.Rows[1][1] + "\t" + "매출합계: " + easyZC03.Rows[1][2] + "\t" + "입금합계: " + easyZC03.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC03.Rows[1][0] + "\t" + "매출건수: " + easyZC03.Rows[1][1] + "\t" + "매출합계: " + easyZC03.Rows[1][2] + "\t" + "입금합계: " + easyZC03.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC03.Rows[2][0] + "\t" + "매출건수: " + easyZC03.Rows[2][1] + "\t" + "매출합계: " + easyZC03.Rows[2][2] + "\t" + "입금합계: " + easyZC03.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC03.Rows[2][0] + "\t" + "매출건수: " + easyZC03.Rows[2][1] + "\t" + "매출합계: " + easyZC03.Rows[2][2] + "\t" + "입금합계: " + easyZC03.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 02 이지 + ZC04 국민카드 (2)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[2]"); // 가맹점그룹 - 영풍문고(이지)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[2]"); // 카드사 - KB카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(이지) + 국민카드 *****");
                WriteLog("***** 영풍문고(이지) + 국민카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC04.Rows[0][0] + "\t" + "매출건수: " + easyZC04.Rows[0][1] + "\t" + "매출합계: " + easyZC04.Rows[0][2] + "\t" + "입금합계: " + easyZC04.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC04.Rows[0][0] + "\t" + "매출건수: " + easyZC04.Rows[0][1] + "\t" + "매출합계: " + easyZC04.Rows[0][2] + "\t" + "입금합계: " + easyZC04.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC04.Rows[1][0] + "\t" + "매출건수: " + easyZC04.Rows[1][1] + "\t" + "매출합계: " + easyZC04.Rows[1][2] + "\t" + "입금합계: " + easyZC04.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC04.Rows[1][0] + "\t" + "매출건수: " + easyZC04.Rows[1][1] + "\t" + "매출합계: " + easyZC04.Rows[1][2] + "\t" + "입금합계: " + easyZC04.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC04.Rows[2][0] + "\t" + "매출건수: " + easyZC04.Rows[2][1] + "\t" + "매출합계: " + easyZC04.Rows[2][2] + "\t" + "입금합계: " + easyZC04.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC04.Rows[2][0] + "\t" + "매출건수: " + easyZC04.Rows[2][1] + "\t" + "매출합계: " + easyZC04.Rows[2][2] + "\t" + "입금합계: " + easyZC04.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 02 이지 + ZC05 삼성카드 (7)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[2]"); // 가맹점그룹 - 영풍문고(이지)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[7]"); // 카드사 - 삼성카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(이지) + 삼성카드 *****");
                WriteLog("***** 영풍문고(이지) + 삼성카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC05.Rows[0][0] + "\t" + "매출건수: " + easyZC05.Rows[0][1] + "\t" + "매출합계: " + easyZC05.Rows[0][2] + "\t" + "입금합계: " + easyZC05.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC05.Rows[0][0] + "\t" + "매출건수: " + easyZC05.Rows[0][1] + "\t" + "매출합계: " + easyZC05.Rows[0][2] + "\t" + "입금합계: " + easyZC05.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC05.Rows[1][0] + "\t" + "매출건수: " + easyZC05.Rows[1][1] + "\t" + "매출합계: " + easyZC05.Rows[1][2] + "\t" + "입금합계: " + easyZC05.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC05.Rows[1][0] + "\t" + "매출건수: " + easyZC05.Rows[1][1] + "\t" + "매출합계: " + easyZC05.Rows[1][2] + "\t" + "입금합계: " + easyZC05.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC05.Rows[2][0] + "\t" + "매출건수: " + easyZC05.Rows[2][1] + "\t" + "매출합계: " + easyZC05.Rows[2][2] + "\t" + "입금합계: " + easyZC05.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC05.Rows[2][0] + "\t" + "매출건수: " + easyZC05.Rows[2][1] + "\t" + "매출합계: " + easyZC05.Rows[2][2] + "\t" + "입금합계: " + easyZC05.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 02 이지 + ZC07 현대카드 (6)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[2]"); // 가맹점그룹 - 영풍문고(이지)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[6]"); // 카드사 - 현대카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(이지) + 현대카드 *****");
                WriteLog("***** 영풍문고(이지) + 현대카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC07.Rows[0][0] + "\t" + "매출건수: " + easyZC07.Rows[0][1] + "\t" + "매출합계: " + easyZC07.Rows[0][2] + "\t" + "입금합계: " + easyZC07.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC07.Rows[0][0] + "\t" + "매출건수: " + easyZC07.Rows[0][1] + "\t" + "매출합계: " + easyZC07.Rows[0][2] + "\t" + "입금합계: " + easyZC07.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC07.Rows[1][0] + "\t" + "매출건수: " + easyZC07.Rows[1][1] + "\t" + "매출합계: " + easyZC07.Rows[1][2] + "\t" + "입금합계: " + easyZC07.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC07.Rows[1][0] + "\t" + "매출건수: " + easyZC07.Rows[1][1] + "\t" + "매출합계: " + easyZC07.Rows[1][2] + "\t" + "입금합계: " + easyZC07.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC07.Rows[2][0] + "\t" + "매출건수: " + easyZC07.Rows[2][1] + "\t" + "매출합계: " + easyZC07.Rows[2][2] + "\t" + "입금합계: " + easyZC07.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC07.Rows[2][0] + "\t" + "매출건수: " + easyZC07.Rows[2][1] + "\t" + "매출합계: " + easyZC07.Rows[2][2] + "\t" + "입금합계: " + easyZC07.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 02 이지 + ZC08 롯데카드 (5)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[2]"); // 가맹점그룹 - 영풍문고(이지)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[5]"); // 카드사 - 롯데카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(이지) + 롯데카드 *****");
                WriteLog("***** 영풍문고(이지) + 롯데카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC08.Rows[0][0] + "\t" + "매출건수: " + easyZC08.Rows[0][1] + "\t" + "매출합계: " + easyZC08.Rows[0][2] + "\t" + "입금합계: " + easyZC08.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC08.Rows[0][0] + "\t" + "매출건수: " + easyZC08.Rows[0][1] + "\t" + "매출합계: " + easyZC08.Rows[0][2] + "\t" + "입금합계: " + easyZC08.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC08.Rows[1][0] + "\t" + "매출건수: " + easyZC08.Rows[1][1] + "\t" + "매출합계: " + easyZC08.Rows[1][2] + "\t" + "입금합계: " + easyZC08.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC08.Rows[1][0] + "\t" + "매출건수: " + easyZC08.Rows[1][1] + "\t" + "매출합계: " + easyZC08.Rows[1][2] + "\t" + "입금합계: " + easyZC08.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                easyZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + easyZC08.Rows[2][0] + "\t" + "매출건수: " + easyZC08.Rows[2][1] + "\t" + "매출합계: " + easyZC08.Rows[2][2] + "\t" + "입금합계: " + easyZC08.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + easyZC08.Rows[2][0] + "\t" + "매출건수: " + easyZC08.Rows[2][1] + "\t" + "매출합계: " + easyZC08.Rows[2][2] + "\t" + "입금합계: " + easyZC08.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion


            #region 03 상품권 + ZC01 비씨카드 (4)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[4]"); // 가맹점그룹 - 영풍문고(상품권)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[4]"); // 카드사 - 비씨카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(상품권) + 비씨카드 *****");
                WriteLog("***** 영풍문고(상품권) + 비씨카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC01.Rows[0][0] + "\t" + "매출건수: " + sangZC01.Rows[0][1] + "\t" + "매출합계: " + sangZC01.Rows[0][2] + "\t" + "입금합계: " + sangZC01.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC01.Rows[0][0] + "\t" + "매출건수: " + sangZC01.Rows[0][1] + "\t" + "매출합계: " + sangZC01.Rows[0][2] + "\t" + "입금합계: " + sangZC01.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC01.Rows[1][0] + "\t" + "매출건수: " + sangZC01.Rows[1][1] + "\t" + "매출합계: " + sangZC01.Rows[1][2] + "\t" + "입금합계: " + sangZC01.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC01.Rows[1][0] + "\t" + "매출건수: " + sangZC01.Rows[1][1] + "\t" + "매출합계: " + sangZC01.Rows[1][2] + "\t" + "입금합계: " + sangZC01.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC01.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC01.Rows[2][0] + "\t" + "매출건수: " + sangZC01.Rows[2][1] + "\t" + "매출합계: " + sangZC01.Rows[2][2] + "\t" + "입금합계: " + sangZC01.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC01.Rows[2][0] + "\t" + "매출건수: " + sangZC01.Rows[2][1] + "\t" + "매출합계: " + sangZC01.Rows[2][2] + "\t" + "입금합계: " + sangZC01.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 03 상품권 + ZC02 하나카드 (10)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[4]"); // 가맹점그룹 - 영풍문고(상품권)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[10]"); // 카드사 - 하나카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(상품권) + 하나카드 *****");
                WriteLog("***** 영풍문고(상품권) + 하나카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC02.Rows[0][0] + "\t" + "매출건수: " + sangZC02.Rows[0][1] + "\t" + "매출합계: " + sangZC02.Rows[0][2] + "\t" + "입금합계: " + sangZC02.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC02.Rows[0][0] + "\t" + "매출건수: " + sangZC02.Rows[0][1] + "\t" + "매출합계: " + sangZC02.Rows[0][2] + "\t" + "입금합계: " + sangZC02.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC02.Rows[1][0] + "\t" + "매출건수: " + sangZC02.Rows[1][1] + "\t" + "매출합계: " + sangZC02.Rows[1][2] + "\t" + "입금합계: " + sangZC02.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC02.Rows[1][0] + "\t" + "매출건수: " + sangZC02.Rows[1][1] + "\t" + "매출합계: " + sangZC02.Rows[1][2] + "\t" + "입금합계: " + sangZC02.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC02.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC02.Rows[2][0] + "\t" + "매출건수: " + sangZC02.Rows[2][1] + "\t" + "매출합계: " + sangZC02.Rows[2][2] + "\t" + "입금합계: " + sangZC02.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC02.Rows[2][0] + "\t" + "매출건수: " + sangZC02.Rows[2][1] + "\t" + "매출합계: " + sangZC02.Rows[2][2] + "\t" + "입금합계: " + sangZC02.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 03 상품권 + ZC03 신한카드 (3)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[4]"); // 가맹점그룹 - 영풍문고(상품권)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[3]"); // 카드사 - 신한카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(상품권) + 신한카드 *****");
                WriteLog("***** 영풍문고(상품권) + 신한카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC03.Rows[0][0] + "\t" + "매출건수: " + sangZC03.Rows[0][1] + "\t" + "매출합계: " + sangZC03.Rows[0][2] + "\t" + "입금합계: " + sangZC03.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC03.Rows[0][0] + "\t" + "매출건수: " + sangZC03.Rows[0][1] + "\t" + "매출합계: " + sangZC03.Rows[0][2] + "\t" + "입금합계: " + sangZC03.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC03.Rows[1][0] + "\t" + "매출건수: " + sangZC03.Rows[1][1] + "\t" + "매출합계: " + sangZC03.Rows[1][2] + "\t" + "입금합계: " + sangZC03.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC03.Rows[1][0] + "\t" + "매출건수: " + sangZC03.Rows[1][1] + "\t" + "매출합계: " + sangZC03.Rows[1][2] + "\t" + "입금합계: " + sangZC03.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC03.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC03.Rows[2][0] + "\t" + "매출건수: " + sangZC03.Rows[2][1] + "\t" + "매출합계: " + sangZC03.Rows[2][2] + "\t" + "입금합계: " + sangZC03.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC03.Rows[2][0] + "\t" + "매출건수: " + sangZC03.Rows[2][1] + "\t" + "매출합계: " + sangZC03.Rows[2][2] + "\t" + "입금합계: " + sangZC03.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 03 상품권 + ZC04 국민카드 (2)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[4]"); // 가맹점그룹 - 영풍문고(상품권)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[2]"); // 카드사 - KB카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(상품권) + 국민카드 *****");
                WriteLog("***** 영풍문고(상품권) + 국민카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC04.Rows[0][0] + "\t" + "매출건수: " + sangZC04.Rows[0][1] + "\t" + "매출합계: " + sangZC04.Rows[0][2] + "\t" + "입금합계: " + sangZC04.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC04.Rows[0][0] + "\t" + "매출건수: " + sangZC04.Rows[0][1] + "\t" + "매출합계: " + sangZC04.Rows[0][2] + "\t" + "입금합계: " + sangZC04.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC04.Rows[1][0] + "\t" + "매출건수: " + sangZC04.Rows[1][1] + "\t" + "매출합계: " + sangZC04.Rows[1][2] + "\t" + "입금합계: " + sangZC04.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC04.Rows[1][0] + "\t" + "매출건수: " + sangZC04.Rows[1][1] + "\t" + "매출합계: " + sangZC04.Rows[1][2] + "\t" + "입금합계: " + sangZC04.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC04.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC04.Rows[2][0] + "\t" + "매출건수: " + sangZC04.Rows[2][1] + "\t" + "매출합계: " + sangZC04.Rows[2][2] + "\t" + "입금합계: " + sangZC04.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC04.Rows[2][0] + "\t" + "매출건수: " + sangZC04.Rows[2][1] + "\t" + "매출합계: " + sangZC04.Rows[2][2] + "\t" + "입금합계: " + sangZC04.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 03 상품권 + ZC05 삼성카드 (7)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[4]"); // 가맹점그룹 - 영풍문고(상품권)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[7]"); // 카드사 - 삼성카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(상품권) + 삼성카드 *****");
                WriteLog("***** 영풍문고(상품권) + 삼성카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC05.Rows[0][0] + "\t" + "매출건수: " + sangZC05.Rows[0][1] + "\t" + "매출합계: " + sangZC05.Rows[0][2] + "\t" + "입금합계: " + sangZC05.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC05.Rows[0][0] + "\t" + "매출건수: " + sangZC05.Rows[0][1] + "\t" + "매출합계: " + sangZC05.Rows[0][2] + "\t" + "입금합계: " + sangZC05.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC05.Rows[1][0] + "\t" + "매출건수: " + sangZC05.Rows[1][1] + "\t" + "매출합계: " + sangZC05.Rows[1][2] + "\t" + "입금합계: " + sangZC05.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC05.Rows[1][0] + "\t" + "매출건수: " + sangZC05.Rows[1][1] + "\t" + "매출합계: " + sangZC05.Rows[1][2] + "\t" + "입금합계: " + sangZC05.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC05.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC05.Rows[2][0] + "\t" + "매출건수: " + sangZC05.Rows[2][1] + "\t" + "매출합계: " + sangZC05.Rows[2][2] + "\t" + "입금합계: " + sangZC05.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC05.Rows[2][0] + "\t" + "매출건수: " + sangZC05.Rows[2][1] + "\t" + "매출합계: " + sangZC05.Rows[2][2] + "\t" + "입금합계: " + sangZC05.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 03 상품권 + ZC07 현대카드 (6)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[4]"); // 가맹점그룹 - 영풍문고(상품권)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[6]"); // 카드사 - 현대카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(상품권) + 현대카드 *****");
                WriteLog("***** 영풍문고(상품권) + 현대카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC07.Rows[0][0] + "\t" + "매출건수: " + sangZC07.Rows[0][1] + "\t" + "매출합계: " + sangZC07.Rows[0][2] + "\t" + "입금합계: " + sangZC07.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC07.Rows[0][0] + "\t" + "매출건수: " + sangZC07.Rows[0][1] + "\t" + "매출합계: " + sangZC07.Rows[0][2] + "\t" + "입금합계: " + sangZC07.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC07.Rows[1][0] + "\t" + "매출건수: " + sangZC07.Rows[1][1] + "\t" + "매출합계: " + sangZC07.Rows[1][2] + "\t" + "입금합계: " + sangZC07.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC07.Rows[1][0] + "\t" + "매출건수: " + sangZC07.Rows[1][1] + "\t" + "매출합계: " + sangZC07.Rows[1][2] + "\t" + "입금합계: " + sangZC07.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC07.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC07.Rows[2][0] + "\t" + "매출건수: " + sangZC07.Rows[2][1] + "\t" + "매출합계: " + sangZC07.Rows[2][2] + "\t" + "입금합계: " + sangZC07.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC07.Rows[2][0] + "\t" + "매출건수: " + sangZC07.Rows[2][1] + "\t" + "매출합계: " + sangZC07.Rows[2][2] + "\t" + "입금합계: " + sangZC07.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            #region 03 상품권 + ZC08 롯데카드 (5)
            element = driver.FindElementByXPath("//*[@id='merGrpId']/option[4]"); // 가맹점그룹 - 영풍문고(상품권)
            element.Click();
            element = driver.FindElementByXPath("//*[@id='cardCo']/option[5]"); // 카드사 - 롯데카드
            element.Click();
            element = driver.FindElementById("searchBtn"); // 조회버튼
            element.Click();
            Thread.Sleep(3000);
            try
            {
                Console.WriteLine("***** 영풍문고(상품권) + 롯데카드 *****");
                WriteLog("***** 영풍문고(상품권) + 롯데카드 *****");
                // 조회결과 첫번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[2]"); // 조회결과 첫번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[3]"); // 조회결과 첫번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[4]"); // 조회결과 첫번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[1]/td[5]"); // 조회결과 첫번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC08.Rows[0][0] + "\t" + "매출건수: " + sangZC08.Rows[0][1] + "\t" + "매출합계: " + sangZC08.Rows[0][2] + "\t" + "입금합계: " + sangZC08.Rows[0][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC08.Rows[0][0] + "\t" + "매출건수: " + sangZC08.Rows[0][1] + "\t" + "매출합계: " + sangZC08.Rows[0][2] + "\t" + "입금합계: " + sangZC08.Rows[0][3]);
                // 조회결과 두번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[2]"); // 조회결과 두번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[3]"); // 조회결과 두번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[4]"); // 조회결과 두번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[2]/td[5]"); // 조회결과 두번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC08.Rows[1][0] + "\t" + "매출건수: " + sangZC08.Rows[1][1] + "\t" + "매출합계: " + sangZC08.Rows[1][2] + "\t" + "입금합계: " + sangZC08.Rows[1][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC08.Rows[1][0] + "\t" + "매출건수: " + sangZC08.Rows[1][1] + "\t" + "매출합계: " + sangZC08.Rows[1][2] + "\t" + "입금합계: " + sangZC08.Rows[1][3]);
                // 조회결과 세번째줄
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[2]"); // 조회결과 세번째줄 - 입금일자
                dr[0] = element.Text.Replace("-", "");
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[3]"); // 조회결과 세번째줄 - 매출건수
                dr[1] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[4]"); // 조회결과 세번째줄 - 매출합계
                dr[2] = element.Text;
                element = driver.FindElementByXPath("//*[@id='list-content']/tr[3]/td[5]"); // 조회결과 세번째줄 - 입금합계
                dr[3] = element.Text;
                sangZC08.Rows.Add(dr);
                Console.WriteLine("입금일자: " + sangZC08.Rows[2][0] + "\t" + "매출건수: " + sangZC08.Rows[2][1] + "\t" + "매출합계: " + sangZC08.Rows[2][2] + "\t" + "입금합계: " + sangZC08.Rows[2][3]); // 조회결과 출력
                WriteLog("입금일자: " + sangZC08.Rows[2][0] + "\t" + "매출건수: " + sangZC08.Rows[2][1] + "\t" + "매출합계: " + sangZC08.Rows[2][2] + "\t" + "입금합계: " + sangZC08.Rows[2][3]);
            }
            catch
            {
                // 아무것도 안함
            }
            #endregion

            WriteLog(_targetDate + " 입금내역 파싱 완료");
            driver.Quit(); // ChromeDriver 종료
        }

        public static void WriteLog(string _message)
        {
            string filePath = Directory.GetCurrentDirectory() + @"\Logs\" + DateTime.Today.ToString("yyyyMMdd") + ".log";
            string dirPath = Directory.GetCurrentDirectory() + @"\Logs";
            string temp;

            DirectoryInfo dirInfo = new DirectoryInfo(dirPath);
            FileInfo fileInfo = new FileInfo(filePath);

            try
            {
                if (dirInfo.Exists != true)
                {
                    Directory.CreateDirectory(dirPath);
                }

                if (fileInfo.Exists != true)
                {
                    using (StreamWriter sw = new StreamWriter(filePath))
                    {
                        temp = string.Format("[{0}] {1}", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"), _message);
                        sw.WriteLine(temp);
                        sw.Close();
                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(filePath))
                    {
                        temp = string.Format("[{0}] {1}", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"), _message);
                        sw.WriteLine(temp);
                        sw.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public static bool RFCInfoSet(string _rfcName, string _rfcFunction)
        {
            try
            {
                // RFC 정보 셋팅
                RfcConfigParameters rfc = new RfcConfigParameters();
                rfc[RfcConfigParameters.Name] = "*";
                rfc[RfcConfigParameters.PoolSize] = Convert.ToString(190);
                rfc[RfcConfigParameters.PeakConnectionsLimit] = Convert.ToString(200);
                rfc[RfcConfigParameters.MaxPoolWaitTime] = Convert.ToString(1);
                rfc[RfcConfigParameters.Client] = "*";
                rfc[RfcConfigParameters.AppServerHost] = "*.*.*.*";
                rfc[RfcConfigParameters.User] = "*";
                rfc[RfcConfigParameters.Password] = "*";
                rfc[RfcConfigParameters.Language] = "KO";
                rfc[RfcConfigParameters.SystemNumber] = "1";
                rfc[RfcConfigParameters.Trace] = "0";
                m_rfcDestination = RfcDestinationManager.GetDestination(rfc);
                m_rfcRepository = m_rfcDestination.Repository;
                m_rfcFunction = m_rfcRepository.CreateFunction(_rfcFunction);
                return true;
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
                return false;
            }
        }

        public static bool InsertRfcDataImport(string _PramName1, string _ParamValue1, string _PramName2, string _ParamValue2)
        {
            try
            {
                m_rfcFunction.SetValue(_PramName1, _ParamValue1);
                m_rfcFunction.SetValue(_PramName2, _ParamValue2);
                return true;
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
                return false;
            }
        }

        public static bool RfcFunctionCall()
        {
            bool l_bResult = false;

            try
            {
                m_rfcFunction.Invoke(m_rfcDestination);
                l_bResult = true;
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);

                string l_strErrorMessage = string.Empty;
                l_strErrorMessage = ex.Message;
                l_bResult = false;
            }
            return l_bResult;
        }

        public static bool InsertRfcDataRows(string _strStructureName, string _strTableName, DataTable _dt)
        {
            try
            {
                RfcStructureMetadata metaData = m_rfcDestination.Repository.GetStructureMetadata(_strStructureName);

                IRfcTable table = m_rfcFunction.GetTable(_strTableName);

                for (int i = 0; i < _dt.Rows.Count; i++)
                {
                    IRfcStructure structData = metaData.CreateStructure();

                    for (int j = 0; j < _dt.Columns.Count; j++)
                    {
                        string l_strColumnName = _dt.Columns[j].ColumnName;

                        structData.SetValue(l_strColumnName, _dt.Rows[i][l_strColumnName]);
                    }
                    table.Append(structData);
                }

                m_rfcFunction.SetValue(_strTableName, table);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
                return false;
            }

            return true;
        }
    }
}
