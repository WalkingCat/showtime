using nietras.SeparatedValues;
var http = new HttpClient();

var events = new List<Event>();
var reader = Sep.Reader(options => options with { Unescape = true });
using var eventsData = reader.FromFile("../data/events.csv");
foreach (var row in eventsData)
{
    var id = row["Id"].ToString();
    if (!string.IsNullOrEmpty(id))
    {
        var evt = new Event
        {
            Id = id,
            Featured = row["Featured"].ToString() == "✅",
            Title = row["Title"].ToString(),
            Region = row["Region"].ToString(),
            Logo = row["Logo"].ToString()
        };
        events.Add(evt);
    }
}

var videos = new Dictionary<string, List<Video>>();
using var videosData = reader.FromFile("../data/videos.csv");
foreach (var row in videosData)
{
    var asf = row["ASF"].ToString();
    var stWmv = row["STWMV"].ToString();
    var slWmv = row["SLWMV"].ToString();
    if (!string.IsNullOrEmpty(asf) || !string.IsNullOrEmpty(stWmv) || !string.IsNullOrEmpty(slWmv))
    {
        var video = new Video
        {
            Id = row["VID"].ToString(),
            EId = row["EID"].ToString(),
            Title = row["Title"].ToString(),
            EventId = row["EVTID"].ToString(),
            Speakers = row["Speakers"].ToString(),
            Asf = asf,
            StWmv = stWmv,
            StSlidesZip = row["STPPT"].ToString() == "✅",
            StSlides = row["STSLIDES"].ToString() == "✅",
            SlId = row["SLID"].ToString(),
            SlWmv = slWmv,
            SlSlidesZip = row["SLPPT"].ToString() == "✅",
            SlSlides = row["SLSLIDES"].ToString() == "✅",
            SlThumbnail = row["SLTN"].ToString() == "✅"
        };
        if (!videos.ContainsKey(video.EventId))
        {
            videos[video.EventId] = new List<Video>();
        }
        videos[video.EventId].Add(video);
    }
}

using var eventsHtml = File.CreateText("../docs/events.html");
eventsHtml.WriteLine("<html><head><link rel='stylesheet' href='styles.css'></head><body class='index'>");
foreach (var evt in events.OrderByDescending(e => e.Featured))
{
    eventsHtml.WriteLine($"<div class='event{(evt.Featured ? " featured" : "")}'><nobr>");
    eventsHtml.WriteLine($"<img src='./images/event_logo_{evt.Id}.png' class='event-logo-thumb{(evt.Featured ? " featured" : "")}{(string.IsNullOrEmpty(evt.Logo) ? " broken" : "")}'/>");
    eventsHtml.WriteLine($"<a href='event_{evt.Id}.html' class='event-title' target='content'>{evt.Title}</a>");
    if (videos.TryGetValue(evt.Id, out var evtVideos))
    {
        eventsHtml.WriteLine($"<span>({evtVideos.Count})</span>");
    }
    if (!string.IsNullOrEmpty(evt.Region))
    {
        eventsHtml.WriteLine($"<span class='event-region'>{evt.Region}</span>");
    }
    eventsHtml.WriteLine("</nobr></div>");
}
eventsHtml.WriteLine("</body></html><body>");

foreach (var evt in events)
{
    Console.WriteLine($"  Event {evt.Id}: {evt.Title}");
    if (!string.IsNullOrEmpty(evt.Logo))
    {
        DownloadFile(evt.Logo, $"../docs/images/event_logo_{evt.Id}.png");
    }

    using var eventHtml = File.CreateText($"../docs/event_{evt.Id}.html");
    eventHtml.WriteLine("<html><head><link rel='stylesheet' href='styles.css'></head><body class='event'>");
    eventHtml.WriteLine("<div class='event-header'>");
    if (!string.IsNullOrEmpty(evt.Logo))
    {
        eventHtml.WriteLine($"<img src='./images/event_logo_{evt.Id}.png' class='event-logo'/>");
    }
    eventHtml.WriteLine($"<div class='event-title-header'>{evt.Title}</div>");
    if (!string.IsNullOrEmpty(evt.Region))
    {
        eventHtml.WriteLine($"<span class='event-region'>Region: {evt.Region}</span>");
    }
    eventHtml.WriteLine("<br/ clear='left'>");
    eventHtml.WriteLine("</div>"); // event-header

    eventHtml.WriteLine("<div class='videos'>");
    if (videos.ContainsKey(evt.Id))
    {
        foreach (var video in videos[evt.Id])
        {
            eventHtml.WriteLine("<div class='video'><hr/>");
            eventHtml.WriteLine($"<div class='video-title'>{(string.IsNullOrEmpty(video.Title) ? "(Unknown Title)" : video.Title)}</div>");
            if (video.SlThumbnail)
            {
                DownloadFile($"https://microsofttech.fr.edgesuite.net/msexp/pictures/spotlight/home/video{video.SlId}.jpg", $"../docs/images/video_thumb_{video.Id}.jpg");
                eventHtml.WriteLine($"<img src='{(video.SlThumbnail ? $"./images/video_thumb_{video.Id}.jpg" : "")}' class='video-thumb{(video.SlThumbnail ? "" : " broken")}'/>");
            }
            if (video.SlSlides)
            {
                DownloadFile($"https://microsofttech.fr.edgesuite.net/msexp/pictures/spotlight/ppt/{video.SlId}/Diapositive1.jpg", $"../docs/images/video_slide_{video.Id}.jpg");
                eventHtml.WriteLine($"<img src='./images/video_slide_{video.Id}.jpg' class='video-slide'>");
            }
            else if (video.StSlides)
            {
                DownloadFile($"https://microsofttech.fr.edgesuite.net/msexp/pictures/images_ppt/{video.Id}/Diapositive1.jpg", $"../docs/images/video_slide_{video.Id}.jpg");
                eventHtml.WriteLine($"<img src='./images/video_slide_{video.Id}.jpg' class='video-slide'>");
            }
            eventHtml.WriteLine($"<span>Speaker(s): {video.Speakers}</span>");

            if (!string.IsNullOrEmpty(video.Asf))
            {
                eventHtml.WriteLine("<div>Streaming: ");
                eventHtml.WriteLine($"<a href='mms://microsofttech.fr.edgesuite.net/msexp/msexp/E{video.EId}/{video.Asf}'>SDVideo</a>(ASF)");
                eventHtml.WriteLine("</div>");
            }

            eventHtml.WriteLine("<div>Downloads: ");
            if (!string.IsNullOrEmpty(video.Asf))
            {
                eventHtml.WriteLine($"<a href='https://microsofttech.fr.edgesuite.net/msexp/msexp/E{video.EId}/{video.Asf}' target='_blank'>SDVideo</a>(ASF)");
            }
            if (!string.IsNullOrEmpty(video.SlWmv))
            {
                eventHtml.WriteLine($"<a href='https://microsofttech.fr.edgesuite.net/msexp/download/spotlight/{video.SlId}/{video.SlId}_{video.SlWmv}.zip' target='_blank'>HDVideo</a>(WMV)");
            }
            else if (!string.IsNullOrEmpty(video.StWmv))
            {
                eventHtml.WriteLine($"<a href='https://microsofttech.fr.edgesuite.net/msexp/download/{video.Id}/{video.Id}_{video.StWmv}.zip' target='_blank'>HDVideo</a>(WMV)");
            }
            if (video.SlSlidesZip)
            {
                eventHtml.WriteLine($"<a href='https://microsofttech.fr.edgesuite.net/msexp/download/spotlight/{video.SlId}/{video.SlId}_pres.zip' target='_blank'>Slides</a><br clear='right'/>");
            }
            else if (video.StSlidesZip)
            {
                eventHtml.WriteLine($"<a href='https://microsofttech.fr.edgesuite.net/msexp/download/{video.Id}/{video.Id}_pres.zip' target='_blank'>Slides</a>");
            }
            eventHtml.WriteLine("</div><br clear='right'/>");
            eventHtml.WriteLine("</div>"); // video
        }
    }
    eventHtml.WriteLine("</div>"); // videos
    eventHtml.WriteLine("</body></html>");
}

void DownloadFile(string uri, string file)
{
    var path = file;
    if (!File.Exists(path))
    {
        bool done = false;
        do
        {
            try
            {
                Console.Write($"  Downloading {file}");
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                var data = http.GetByteArrayAsync(uri).Result;
                if (data.Length > 0)
                {
                    using var fs = File.OpenWrite(path);
                    fs.Write(data);
                    fs.Flush();
                    fs.Close();
                }
                else
                {
                    Console.Write(" [EMPTY]");
                }
                done = true;
            }
            catch (AggregateException ex)
            {
                if (ex.InnerException is HttpRequestException hrex)
                {
                    Console.Write($" [{hrex.StatusCode}]");
                    if ((hrex.StatusCode == System.Net.HttpStatusCode.NotFound)
                        || (hrex.StatusCode == System.Net.HttpStatusCode.GatewayTimeout)
                        || (hrex.StatusCode == System.Net.HttpStatusCode.Conflict))
                    {
                        done = true;
                    }
                }
                else throw;
            }
            catch
            {
                Console.Write(" [ERROR]");
                done = true;
            }
            Console.WriteLine();
        } while (!done);
    }
}

class Event
{
    public required string Id;
    public required bool Featured;
    public required string Title;
    public required string Region;
    public required string Logo;
}

class Video
{
    public required string Id;
    public required string EId;
    public required string Title;
    public required string EventId;
    public required string Speakers;
    public required string Asf;
    public required string StWmv;
    public bool StSlidesZip;
    public bool StSlides;
    public required string SlId;
    public required string SlWmv;
    public bool SlSlidesZip;
    public bool SlThumbnail;
    public bool SlSlides;
}