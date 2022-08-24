using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using Slack.Webhooks;

const string webHookUrl = "https://hooks.slack.com/services/...";

static byte[] Crop(string Img, int Width, int Height, int X, int Y)
{
    //Graphic.SmoothingMode = SmoothingMode.AntiAlias;
    //Graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
    //Graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
    using System.Drawing.Image OriginalImage = System.Drawing.Image.FromFile(Img);
    using System.Drawing.Bitmap bmp = new(Width, Height);
    bmp.SetResolution(OriginalImage.HorizontalResolution, OriginalImage.VerticalResolution);
    using System.Drawing.Graphics Graphic = System.Drawing.Graphics.FromImage(bmp);
    Graphic.DrawImage(OriginalImage, new System.Drawing.Rectangle(0, 0, Width, Height), X, Y, Width, Height, System.Drawing.GraphicsUnit.Pixel);
    var ms = new MemoryStream();
    bmp.Save(ms, OriginalImage.RawFormat);

    return ms.GetBuffer();
}

void PostWebHookToSlack(string text, string caption, string link)
{
    var slackClient = new SlackClient(webHookUrl);
    var slackMessage = new SlackMessage
    {
        Channel = "#gurmans",
        Text = caption,
        IconEmoji = Emoji.MeatOnBone,
        Username = "GurmanBot"
    };
    var slackAttachment = new SlackAttachment
    {
        Fallback = text,
        Text = text,
        Color = "#D00000",
        ImageUrl = link
    };
    slackMessage.Attachments = new List<SlackAttachment> { slackAttachment };
    slackClient.Post(slackMessage);
}

void ScreenShotWebsite(string url, int delay)
{
    var opts = new FirefoxOptions
    {
        AcceptInsecureCertificates = true,
        //in firefox open tab "about:support"
        Profile = new FirefoxProfile(@"C:\Users\ashur\AppData\Roaming\Mozilla\Firefox\Profiles\kmann38c.default-release")
    };
    var driver = new FirefoxDriver(@$"c:\geckodriver", opts);
    driver.Navigate().GoToUrl(url);
    driver.Manage().Window.Size = new System.Drawing.Size(1920, 5000);
    Thread.Sleep(delay);
    var screenshot = (driver as ITakesScreenshot).GetScreenshot();
    screenshot.SaveAsFile("screenshot.png");
    driver.Close();
    driver.Quit();
}

string ProcessNaSolnici(string url)
{
    ScreenShotWebsite(url, 0);
    var path = "solnice-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-fffffff");
    var fs = new FileStream(@$"c:\inetpub\wwwroot\gurmans\{path}.png", FileMode.Create);
    using (var sw = new StreamWriter(fs))
    {
        switch (DateTime.Now.DayOfWeek)
        {
            case DayOfWeek.Monday:
                fs.Write(Crop("screenshot.png", 600, 300, 650, 2500));
                break;
            case DayOfWeek.Tuesday:
                fs.Write(Crop("screenshot.png", 600, 300, 650, 2600));
                break;
            case DayOfWeek.Wednesday:
                fs.Write(Crop("screenshot.png", 600, 300, 650, 2650));
                break;
            case DayOfWeek.Thursday:
                fs.Write(Crop("screenshot.png", 600, 300, 650, 2700));
                break;
            case DayOfWeek.Friday:
                fs.Write(Crop("screenshot.png", 600, 300, 650, 2750));
                break;
        }
    }
    var img = $"http://ashur1337.asuscomm.com/gurmans/{path}.png";
    PostWebHookToSlack(url, "Na Solnici", img);
    return img;
}

string ProcessUCertu(string url)
{
    ScreenShotWebsite(url, 5000);
    var path = "ucertu-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-fffffff");
    var fs = new FileStream(@$"c:\inetpub\wwwroot\gurmans\{path}.png", FileMode.Create);
    using (var sw = new StreamWriter(fs))
    {
        fs.Write(Crop("screenshot.png", 700, 450, 660, 530));
    }
    var img = $"http://ashur1337.asuscomm.com/gurmans/{path}.png";
    PostWebHookToSlack(url, "U certu", img);
    return img;
}

string ProcessOpice(string url)
{
    ScreenShotWebsite(url, 500);
    var path = "opice-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-fffffff");
    var fs = new FileStream(@$"c:\inetpub\wwwroot\gurmans\{path}.png", FileMode.Create);
    using (var sw = new StreamWriter(fs))
    {
        fs.Write(Crop("screenshot.png", 600, 280, 380, 470));
    }
    var img = $"http://ashur1337.asuscomm.com/gurmans/{path}.png";
    PostWebHookToSlack(url, "Pivni Opice", img);
    return img;
}

string ProcessKolkovna(string url)
{
    ScreenShotWebsite(url, 500);
    var path = "kolkovna" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-fffffff");
    var fs = new FileStream(@$"c:\inetpub\wwwroot\gurmans\{path}.png", FileMode.Create);
    using (var sw = new StreamWriter(fs))
    {        
        switch (DateTime.Now.DayOfWeek)
        {
            case DayOfWeek.Monday:
                fs.Write(Crop("screenshot.png", 540, 180, 540, 270));
                break;
            case DayOfWeek.Tuesday:
                fs.Write(Crop("screenshot.png", 540, 180, 540, 450));
                break;
            case DayOfWeek.Wednesday:
                fs.Write(Crop("screenshot.png", 540, 180, 540, 620));
                break;
            case DayOfWeek.Thursday:
                fs.Write(Crop("screenshot.png", 540, 180, 540, 800));
                break;
            case DayOfWeek.Friday:
                fs.Write(Crop("screenshot.png", 540, 180, 540, 980));
                break;
        }
    }
    var img = $"http://ashur1337.asuscomm.com/gurmans/{path}.png";
    PostWebHookToSlack(url, "Kolkovna", img);
    return img;
}

string ProcessUTomana(string url)
{
    ScreenShotWebsite(url, 500);
    var path = "utomana-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss-fffffff");
    var fs = new FileStream(@$"c:\inetpub\wwwroot\gurmans\{path}.png", FileMode.Create);
    using (var sw = new StreamWriter(fs))
    {
        fs.Write(Crop("screenshot.png", 670, 160, 780, 900));
    }
    var img = $"http://ashur1337.asuscomm.com/gurmans/{path}.png";
    PostWebHookToSlack(url, "U Tomana", img);
    return img;
}

string ProcessJakoby(string url)
{
    ScreenShotWebsite(url, 500);
    var path = "jakoby-" + DateTime.Now.ToString("jakoby-yyyy-MM-dd-HH-mm-ss-fffffff");
    var fs = new FileStream(@$"c:\inetpub\wwwroot\gurmans\{path}.png", FileMode.Create);
    using (var sw = new StreamWriter(fs))
    {
        fs.Write(Crop("screenshot.png", 670, 170, 780, 930));
    }
    var img = $"http://ashur1337.asuscomm.com/gurmans/{path}.png";
    PostWebHookToSlack(url, "Jakoby", img);
    return img;
}

string img;
img = ProcessNaSolnici("https://www.nasolnici.cz/");
img = ProcessUCertu("http://ucertu.cz/nabidka-dvorakova/");
img = ProcessOpice("https://www.pivniopice.cz/");
img = ProcessKolkovna("https://www.kolkovna.cz/cs/stopkova-plzenska-pivnice-16/denni-menu");
img = ProcessUTomana("https://www.menicka.cz/6897-restaurace-u-tomana.html");
img = ProcessJakoby("https://www.menicka.cz/2717-restaurace-jakoby.html");




