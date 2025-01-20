
using EAGetMail;
using Svg;
using System.Drawing;
using System.Text.RegularExpressions;
using ZXing.Windows.Compatibility;
using Attachment = EAGetMail.Attachment;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Spire.Pdf.Graphics;
using Spire.Pdf;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml.Packaging;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // Unique URL list
            HashSet<string> uniqueUrls = new HashSet<string>();
            string emlFilePath = "C://mails/Capturing And Analyzing QRCodes in An Email Part 1.eml";
           

            if (string.IsNullOrEmpty(emlFilePath) || !File.Exists(emlFilePath))
            {
                Console.WriteLine("Invalid file path. Exiting.");
                return;
            }

            // Load email from the specified file
            Mail email = new Mail("TryIt");
            byte[] emlContent = await File.ReadAllBytesAsync(emlFilePath);

            email.Load(emlContent);
            Console.WriteLine($"Subject: {email.Subject}");

            // Check mail attachments for QR codes
            foreach (Attachment attachment in email.Attachments)
            {
                string tempFile = string.Empty;
                try
                {
                    // Generate a unique file name using GUID
                    string extension = Path.GetExtension(attachment.Name).ToLower();
                    if (string.IsNullOrEmpty(extension))
                    {
                        // If no extension is found, assume the file is an image (default to .jpg)
                        extension = ".jpg";
                    }

                    // Generate a GUID-based unique file name
                    string uniqueFileName = Guid.NewGuid().ToString() + extension;
                    tempFile = Path.Combine(Path.GetTempPath(), uniqueFileName);

                    // Save the attachment with the unique file name
                    attachment.SaveAs(tempFile, true);
                    Console.WriteLine($"Attachment saved: {tempFile}");

                    // Now check if the saved file is an image
                    if (IsImageFile(tempFile))
                    {
                        string[]? qrCodeContentList = ScanBarcode(tempFile);
                        if (qrCodeContentList != null && qrCodeContentList.Length > 0)
                        {
                            qrCodeContentList.ToList().ForEach(qrCodeContent =>
                            {
                                string extractedUrl = ExtractUrl(qrCodeContent);
                                if (!string.IsNullOrEmpty(extractedUrl) && uniqueUrls.Add(extractedUrl))
                                {
                                    Console.WriteLine($"Unique URL Found: {extractedUrl}");
                                }
                            });
                        }
                    }                                      
                    else if (extension.Trim() == ".pdf")
                    {
                        // Call method for PDF processing
                        ProcessPdfAttachment(tempFile, uniqueUrls);
                    }                   
                    else if (extension.Trim() == ".docx" || extension.Trim() == ".doc")
                    {
                        // Call method for Word processing
                        ProcessWordAttachment(tempFile, uniqueUrls);
                    }
                    else
                    {
                        Console.WriteLine($"Unsupported attachment type: {extension}");
                    }                  
                    Console.WriteLine($"Deleted temporary file: {tempFile}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing attachment: {ex.Message}");
                }
                finally
                {
                    if (File.Exists(tempFile))
                        File.Delete(tempFile);
                }
            }


            // Check HTML body for QR codes
            if (IsHtml(email))
            {
                var imageList = await GetMailImages(email);

                foreach (string imgSrc in imageList)
                {                    
                    string[]? qrCodeContentList = ScanBarcode(imgSrc);
                    if (qrCodeContentList != null && qrCodeContentList.Length > 0)
                    {
                        qrCodeContentList.ToList().ForEach(qrCodeContent =>
                        {
                            string extractedUrl = ExtractUrl(qrCodeContent);
                            if (!string.IsNullOrEmpty(extractedUrl) && uniqueUrls.Add(extractedUrl))
                            {
                                Console.WriteLine($"Unique URL Found in HTML: {extractedUrl}");
                            }
                        });
                    }                    
                }
            }


            // Print all unique URLs
            Console.WriteLine("\nAll Unique URLs:");
            foreach (string url in uniqueUrls)
            {
                Console.WriteLine(url);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static bool IsImageFile(string filePath)
    {
        // Add a check for null or empty file path
        if (string.IsNullOrEmpty(filePath)) return false;

        string[] validExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
        string? extension = Path.GetExtension(filePath)?.ToLower();

        // Ensure the file has an extension
        if (string.IsNullOrEmpty(extension)) return false;

        return Array.Exists(validExtensions, ext => ext == extension);
    }
    static string[]? ScanBarcode(string imagePath)
    {
        try
        {
            // Barcode reader instance
            var barcodeReader = new BarcodeReader();

            // Check if the imagePath is a 'data:' URI (Base64 encoded image)
            if (imagePath.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                // Extract Base64 string after "data:image/...;base64,"
                string base64Data = imagePath.Substring(imagePath.IndexOf(",") + 1);

                // Convert the Base64 string to byte array
                byte[] imageBytes = Convert.FromBase64String(base64Data);

                // Check if the data is SVG
                if (imagePath.Contains("image/svg+xml"))
                {
                    // Load the SVG content from the byte array
                    using (var memoryStream = new MemoryStream(imageBytes))
                    {
                        var svgDocument = SvgDocument.Open<SvgDocument>(memoryStream);

                        // Convert the SVG to a Bitmap
                        using (var bitmap = svgDocument.Draw())
                        {
                            // Now you can scan the barcode from the bitmap         
                            var result = barcodeReader.DecodeMultiple(bitmap);
                            if (result != null && result.Length > 0)
                            {
                                return result.Select(r => r.Text).ToArray();
                            }
                            else
                            {
                                Console.WriteLine("No QR code or barcode detected in the image.");
                                return null;
                            }
                        }
                    }
                }
                else
                {
                    // If not SVG, treat it as another image type (e.g., PNG, JPEG)
                    using (var memoryStream = new MemoryStream(imageBytes))
                    {
                        using (var bitmap = new Bitmap(memoryStream))
                        {
                            var result = barcodeReader.DecodeMultiple(bitmap);
                            if (result != null && result.Length > 0)
                            {
                                return result.Select(r => r.Text).ToArray();
                            }
                            else
                            {
                                Console.WriteLine("No QR code or barcode detected in the image.");
                                return null;
                            }
                        }
                    }
                }
            }
            else if (Uri.IsWellFormedUriString(imagePath, UriKind.Absolute))
            {
                // If it's a regular URL, download it
                using (var httpClient = new HttpClient())
                {
                    byte[] imageBytes = httpClient.GetByteArrayAsync(imagePath).Result;

                    using (var memoryStream = new MemoryStream(imageBytes))
                    {
                        // Load image from MemoryStream
                        using (var bitmap = new Bitmap(memoryStream))
                        {
                            var result = barcodeReader.DecodeMultiple(bitmap);
                            if (result != null && result.Length > 0)
                            {
                                return result.Select(r => r.Text).ToArray();
                            }
                            else
                            {
                                Console.WriteLine("No QR code or barcode detected in the image.");
                                return null;
                            }
                        }
                    }
                }
            }
            else
            {
                // If it's a local file path, load it directly
                using (var bitmap = (Bitmap)Image.FromFile(imagePath))
                {
                    var result = barcodeReader.DecodeMultiple(bitmap);
                    if (result != null && result.Length > 0)
                    {
                        return result.Select(r => r.Text).ToArray();
                    }
                    else
                    {
                        Console.WriteLine("No QR code or barcode detected in the image.");
                        return null;
                    }
                }
            }
        }
        catch (System.FormatException ex)
        {
            Console.WriteLine($"Error decoding Base64 image data: {ex.Message}");
            return null;
        }
        catch (FileNotFoundException)
        {
            Console.WriteLine("Error: The specified image file was not found.");
            return null;
        }
        catch (OutOfMemoryException)
        {
            Console.WriteLine("Error: The file is not a valid image or is too large.");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error scanning barcode: {ex.Message}");
            return null;
        }
    }

    static IEnumerable<string> ExtractImageSources(string html)
    {
        // Updated regular expression to capture both URLs and data URIs in src attributes                
        var matches = Regex.Matches(html, @"<img[^>]+src=['""](data:[^'""]*|[^'""]*)['""]", RegexOptions.IgnoreCase);



        foreach (Match match in matches)
        {
            // Yield each src attribute value
            string src = match.Groups[1].Value;
            if (!string.IsNullOrEmpty(src))
            {
                yield return src;
            }
        }
    }


    public static async Task<List<string>> GetMailImages(Mail email)
    {
        try
        {
            // EML dosyasındaki HTML içeriği al
            string htmlContent = email.HtmlBody;

            if (string.IsNullOrEmpty(htmlContent))
            {
                throw new Exception("HTML content is empty or invalid.");
            }           

            // WebDriverManager          
            var _config = new ChromeConfig();
            //Console.WriteLine($"Chrome Browser Version: {_config.GetMatchingBrowserVersion()}");
            new DriverManager("C://custom//chromedriver//location").SetUpDriver(_config, _config.GetMatchingBrowserVersion());   

            // ChromeDriver başlat
            var options = new ChromeOptions();
            options.AddArgument("--headless");
            options.AddArgument("--disable-gpu");
            //options.AddArgument("--no-sandbox");
  
            options.AddUserProfilePreference("download.prompt_for_download", false); 
            options.AddUserProfilePreference("download.default_directory", "");     
            options.AddUserProfilePreference("profile.default_content_settings.popups", 0);
            options.AddUserProfilePreference("safebrowsing.enabled", true); 
            options.AddUserProfilePreference("profile.default_content_settings.automatic_downloads", 1); 

            //Extra Security
            options.AddArgument("--incognito");
            options.AddArgument("--disable-dev-shm-usage");
            options.AddArgument("--disable-web-security");

            //Only Allow Https.
            options.AddArgument("--deny-permission-prompts");
            options.AddArgument("--block-new-web-contents");

            // Script ve Plugin Setup
            options.AddArgument("--disable-plugins");
            options.AddUserProfilePreference("profile.managed_default_content_settings.javascript", 2); // Disabled JavaScript.

            // Save Data and Cache Setup
            options.AddArgument("--disable-cache");
            options.AddArgument("--disable-application-cache"); 
            options.AddArgument("--disable-logging");

            // Security Policy
            options.AddArgument("--disable-popup-blocking"); 
            options.AddArgument("--disable-extensions");
            options.AddArgument("--disable-blink-features=AutomationControlled"); 
            options.AddArgument("--disable-features=IsolateOrigins,site-per-process");


            //var chromeDriverPath = @"C:\tools\chromedriver\"; // Adjust your path as needed
            using (var driver = new ChromeDriver(options))
            {

                string tempHtmlPath = string.Empty;
                try
                {
                    driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(15); // Sayfa yükleme zamanı
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10); // Eleman bulunması için bekleme

                    // HTML içeriği bir dosyaya yaz
                    string uniqueFileName = Guid.NewGuid().ToString() + ".html";
                    tempHtmlPath = Path.Combine(Path.GetTempPath(), uniqueFileName);

                    //string tempHtmlPath = Path.Combine(Path.GetTempPath(), "tempEmail.html");
                    await File.WriteAllTextAsync(tempHtmlPath, htmlContent);

                    // HTML dosyasını yükle
                    driver.Navigate().GoToUrl($"file:///{tempHtmlPath}");

                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(d => d.FindElements(By.TagName("img")).Any()); // Wait for any images to load

                    // Use JavaScript to extract all image sources
                    var allImages = ((IJavaScriptExecutor)driver).ExecuteScript(
                        "return Array.from(document.querySelectorAll('img')).map(img => img.src);"
                    ) as IReadOnlyCollection<object>;

                    File.Delete(tempHtmlPath);

                    // Convert to a list of strings
                    return allImages?.Select(img => img?.ToString())
                                     .Where(img => img != null)
                                     .Distinct()
                                     .Cast<string>()
                                     .ToList() ?? new List<string>();
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine("Timeout: No images with 'data:image' found.");
                    return new List<string>(); // Boş liste döndür
                }
                finally
                {
                    if (File.Exists(tempHtmlPath))
                        File.Delete(tempHtmlPath);
                    // Close Browser
                    driver.Quit();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting images from HTML: {ex.Message}");
            return new List<string>();
        }
    }
    static void ProcessPdfAttachment(string tempFile, HashSet<string> uniqueUrls)
    {
        String tempPDFImageFile = String.Empty;
        try
        {
            using (PdfDocument pdf = new PdfDocument())
            {
                // Load the PDF document
                pdf.LoadFromFile(tempFile);
                string uniquePDFimageName = Guid.NewGuid().ToString();
                // Loop through each page in the PDF
                for (int i = 0; i < pdf.Pages.Count; i++)
                {
                    try
                    {
                        // Generate a unique file name for each page image
                        string uniqueImageName = $"{uniquePDFimageName}-{i + 1}.png";
                        tempPDFImageFile = Path.Combine(Path.GetTempPath(), uniqueImageName);

                        // Convert each page to an image with the specified DPI
                        Image image = pdf.SaveAsImage(i, PdfImageType.Bitmap, 300, 300);
                        image.Save(tempPDFImageFile, ImageFormat.Png);                      
                        Console.WriteLine($"Saved PDF page as image: {tempPDFImageFile}");

                        string[]? qrCodeContentList = ScanBarcode(tempPDFImageFile);
                        if (qrCodeContentList != null && qrCodeContentList.Length > 0)
                        {
                            qrCodeContentList.ToList().ForEach(qrCodeContent =>
                            {
                                string extractedUrl = ExtractUrl(qrCodeContent);
                                if (!string.IsNullOrEmpty(extractedUrl) && uniqueUrls.Add(extractedUrl))
                                {
                                    Console.WriteLine($"Unique URL Found: {extractedUrl}");
                                }
                            });
                        }

                        File.Delete(tempPDFImageFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing PDF attachment: {ex.Message}");
                    }
                    finally
                    {
                        if (File.Exists(tempPDFImageFile))
                            File.Delete(tempPDFImageFile);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing PDF attachment: {ex.Message}");
        }
        finally
        {
            if (File.Exists(tempPDFImageFile))
                File.Delete(tempPDFImageFile);
        }
    }
    static void ProcessWordAttachment(string tempFile, HashSet<string> uniqueUrls)
    {
        String tempWordImageFile = String.Empty;
        try
        {
            if (string.IsNullOrEmpty(tempFile) || !File.Exists(tempFile))
            {
                Console.WriteLine("Invalid .docx file path.");
                return;
            }
            // Open the .docx || .doc file
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(tempFile, false))
            {
                var mainPart = wordDoc.MainDocumentPart;

                // Check for embedded images in the document
                if (mainPart != null && mainPart.ImageParts.Any())
                {
                    foreach (var imagePart in mainPart.ImageParts)
                    {
                        try
                        {
                            // Save the image to a temporary file for QR code scanning
                            tempWordImageFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                            using (var stream = imagePart.GetStream())
                            using (var fileStream = new FileStream(tempWordImageFile, FileMode.Create, FileAccess.Write))
                            {
                                stream.CopyTo(fileStream);
                            }

                            Console.WriteLine($"Image extracted: {tempWordImageFile}");

                            // Scan the image for QR codes                         
                            string[]? qrCodeContentList = ScanBarcode(tempWordImageFile);
                            if (qrCodeContentList != null && qrCodeContentList.Length > 0)
                            {
                                qrCodeContentList.ToList().ForEach(qrCodeContent =>
                                {
                                    string extractedUrl = ExtractUrl(qrCodeContent);
                                    if (!string.IsNullOrEmpty(extractedUrl) && uniqueUrls.Add(extractedUrl))
                                    {
                                        Console.WriteLine($"Unique URL Found in .docx: {extractedUrl}");
                                    }
                                });
                            }

                            // Clean up the temporary file
                            File.Delete(tempWordImageFile);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing image in .docx: {ex.Message}");
                        }
                        finally
                        {
                            if (File.Exists(tempWordImageFile))
                                File.Delete(tempWordImageFile);
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No images found in the .docx file.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing .docx file: {ex.Message}");
        }
        finally
        {
            if (File.Exists(tempWordImageFile))
                File.Delete(tempWordImageFile);
        }
    }
    static string ExtractUrl(string text)
    {
        var match = Regex.Match(text, @"https?://[\w./?=&%-]+", RegexOptions.IgnoreCase);
        return match.Success ? match.Value : null;
    }

    static bool IsHtml(Mail email)
    {
        return !string.IsNullOrEmpty(email.HtmlBody);
    }
}
