using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using csExWB;
using HtmlAgilityPack;
using IfacesEnumsStructsClasses;

namespace WindowsFormsBrowser
{
    public partial class Form1 : Form
    {
        private cEXWB cEXWB1;
        string data;
        int i = 1;
        int j = 1;
        public string Address = string.Empty;
        public string Name = string.Empty;
        public string PhoneNumber = string.Empty;
        public string Email = string.Empty;
        public string website = string.Empty;
        public string location = Application.StartupPath + "\\LogData.txt";
        public Form1()
        {
            this.cEXWB1 = new csExWB.cEXWB();
            this.cEXWB1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
| System.Windows.Forms.AnchorStyles.Left)
| System.Windows.Forms.AnchorStyles.Right)));
            this.cEXWB1.Border3DEnabled = false;
            this.cEXWB1.DocumentSource = "<HTML><HEAD></HEAD>\r\n<BODY></BODY></HTML>";
            this.cEXWB1.DocumentTitle = "";
            this.cEXWB1.DownloadActiveX = true;
            this.cEXWB1.DownloadFrames = true;
            this.cEXWB1.DownloadImages = true;
            this.cEXWB1.DownloadJava = true;
            this.cEXWB1.DownloadScripts = true;
            this.cEXWB1.DownloadSounds = true;
            this.cEXWB1.DownloadVideo = true;
            this.cEXWB1.FileDownloadDirectory = "C:\\DEV\\";
            this.cEXWB1.Location = new System.Drawing.Point(4, 212);
            this.cEXWB1.LocationUrl = "about:blank";
            this.cEXWB1.Name = "cEXWB1";
            this.cEXWB1.ObjectForScripting = null;
            this.cEXWB1.OffLine = false;
            this.cEXWB1.RegisterAsBrowser = false;
            this.cEXWB1.RegisterAsDropTarget = false;
            this.cEXWB1.RegisterForInternalDragDrop = true;
            this.cEXWB1.ScrollBarsEnabled = true;
            this.cEXWB1.SendSourceOnDocumentCompleteWBEx = false;
            this.cEXWB1.Silent = false;
            this.cEXWB1.Size = new System.Drawing.Size(1000, 528);
            this.cEXWB1.TabIndex = 25;
            this.cEXWB1.Text = "cEXWB1";
            this.cEXWB1.TextSize = IfacesEnumsStructsClasses.TextSizeWB.Medium;
            this.cEXWB1.UseInternalDownloadManager = true;
            this.cEXWB1.WBDOCDOWNLOADCTLFLAG = 112;
            this.cEXWB1.WBDOCHOSTUIDBLCLK = IfacesEnumsStructsClasses.DOCHOSTUIDBLCLK.DEFAULT;
            this.cEXWB1.WBDOCHOSTUIFLAG = 262276;
            this.cEXWB1.DocumentComplete += new csExWB.DocumentCompleteEventHandler(this.cEXWB1_DocumentComplete);
            
            InitializeComponent();
        }
        private void cEXWB1_DocumentComplete(object sender, DocumentCompleteEventArgs e)
        {

            //Your Local Australian Business Directory
            if(j==29)
            {
                MessageBox.Show("Page is 29 completed");
            }
            if ((i==1) && (this.cEXWB1.DocumentTitle.ToString().Contains("Your Local Australian Business Directory")))
            {
                this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Listing+Services&locationClue=All+States&pageNumber="+j+"&referredBy=www.yellowpages.com.au&&eventType=pagination");
               // this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Sales+Advisory+Services&locationClue=All+States&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&eventType=pagination");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Development&locationClue=All+States&pageNumber="+j+"&referredBy=www.yellowpages.com.au&&eventType=pagination");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Auctioneers&locationClue=All+States&pageNumber="+j+"&referredBy=www.yellowpages.com.au&&eventType=pagination");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=commercial+real+estate+agents&locationClue=All+States&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&eventType=pagination");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Agents&eventType=pagination&locationClue=All+States&openNow=false&pageNumber="+j+"&referredBy=www.yellowpages.com.au&&state=NSW");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Truline+Australia+Pty+Ltd&locationClue=All+States&lat=&lon=&selectedViewMode=list");
                data = this.cEXWB1.DocumentSource.ToString();
                i++;
            } else if (i == 2)
            {
                #region comment2
                //IfacesEnumsStructsClasses.IHTMLElementCollection data = this.cEXWB1.GetElementsByTagName(false,"div");
                //foreach (IHTMLElement2 ce in data)
                //{
                //    IHTMLElement cee = (IHTMLElement)ce;
                //    // object obj = cee.getAttribute("href", 1);
                //    //collapser__link collapser__link--dense
                //   // object obj = cee.getAttribute("class", 1);

                //     if (cee.outerHTML.ToString().Contains("search-results search-results-data listing-group"))
                //    {
                //      //  cee.click();
                //        // break;
                //    }

                //    //MessageBox.Show(cee.innerHTML);
                //}
                #endregion
                HtmlAgilityPack.HtmlDocument h1 = new HtmlAgilityPack.HtmlDocument();
                h1.LoadHtml(this.cEXWB1.DocumentSource.ToString());
                var datas = h1.DocumentNode.SelectSingleNode("//html/body/div[1]/div/div[3]/div/div/div[2]/div/div[2]/div[2]/div");
                //class="cell in-area-cell find-show-more-trial   middle-cell"
                var filter1 = datas.SelectSingleNode("//div[@class='cell in-area-cell find-show-more-trial   middle-cell']");
                foreach(var Filter in datas.SelectNodes("//div[@class='cell in-area-cell find-show-more-trial   middle-cell']"))
                {

                    var items= Filter;
                    for (int k = 0; k < items.SelectNodes("//a[@class='listing-name']").Count; k++)
                    {
                        Address = string.Empty;
                        Name = string.Empty;
                        PhoneNumber = string.Empty;
                        Email = string.Empty;
                        website = string.Empty;
                        #region comment1
                        ////class="listing-name"
                        //foreach (var item2 in items.SelectNodes("//a"))
                        //{
                        //}
                        //<span class="contact-text">(02) 9868 3333</span>
                        //class="contact contact-main contact-url " website get href value
                        //class="contact contact-main contact-email " get data-email=
                        #endregion

                        try
                        {
                            Name = items.SelectNodes("//a[@class='listing-name']")[k].InnerHtml.ToString();
                        }
                        catch (Exception ex)
                        {
                            Name = "";
                        }

                        try
                        {
                            Address = items.SelectNodes("//p[@class='listing-address mappable-address mappable-address-with-poi']")[k].InnerHtml.ToString();
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                Address = items.SelectNodes("//p[@class='listing-address mappable-address']")[k].InnerHtml.ToString();
                            }
                            catch (Exception ex1)
                            {
                                try
                                {
                                    Address = items.SelectNodes("//p[@class='listing-heading']")[k].InnerText.ToString();
                                }
                                catch (Exception ex2)
                                {
                                    Address = "";
                                }
                            }
                        }

                        try
                        {
                            PhoneNumber = items.SelectNodes("//a[@class='click-to-call contact contact-preferred contact-phone ']")[k].InnerText.ToString().Replace("\n","").Trim();
                        }
                        catch (Exception ex)
                        {
                            PhoneNumber = "";
                        }

                        try
                        {
                            website = items.SelectNodes("//a[@class='contact contact-main contact-url ']")[k].GetAttributeValue("href", "").ToString();
                        }
                        catch (Exception ex)
                        {
                            website = "";
                        }
                        try
                        {
                            Email = items.SelectNodes("//a[@class='contact contact-main contact-email ']")[k].GetAttributeValue("data-email", "").ToString();
                        }
                        catch (Exception ex)
                        {
                            Email = "";
                        }
                        StringBuilder str = new StringBuilder();
                        LogData(Name + "\t" + Address + "\t" + PhoneNumber + "\t" + website + "\t" + Email);
                    }
                    break;
                }
                Console.WriteLine(j + " pages completed");
                j++;
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Development&locationClue=All+States&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&eventType=pagination");
                // this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Agents&eventType=pagination&locationClue=All+States&openNow=false&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&state=NSW");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=commercial+real+estate+agents&locationClue=All+States&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&eventType=pagination");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Auctioneers&locationClue=All+States&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&eventType=pagination");
                //this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Sales+Advisory+Services&locationClue=All+States&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&eventType=pagination");
                this.cEXWB1.Navigate2("https://www.yellowpages.com.au/search/listings?clue=Real+Estate+Listing+Services&locationClue=All+States&pageNumber=" + j + "&referredBy=www.yellowpages.com.au&&eventType=pagination");
            }
        }
        public void TestData()
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.cEXWB1.Navigate2("https://www.yellowpages.com.au/");
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
        public void LogData(string LogData)
        {
            
            if(!File.Exists(location.ToString()))
            {
                File.Create(Location.ToString());
            }
            FileStream Exceptionstream = new FileStream(location, FileMode.Append);
            StreamWriter Exceptionwriter = new StreamWriter(Exceptionstream);
            Exceptionwriter.WriteLine(LogData);
            Exceptionwriter.Close();
            Exceptionstream.Close();

        }
    }
}
