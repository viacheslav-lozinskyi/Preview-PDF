using PdfiumViewer;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace resource.preview
{
    internal class VSPreview : extension.AnyPreview
    {
        protected override void _Execute(atom.Trace context, int level, string url, string file)
        {
            var a_Name = GetUrlPreview(file, ".png");
            {
                context.
                    SetAlignment(NAME.ALIGNMENT.TOP).
                    SetFontState(NAME.FONT_STATE.BLINK).
                    SetProgress(NAME.PROGRESS.INFINITE).
                    SetUrlPreview(a_Name).
                    SendPreview(NAME.TYPE.INFO, url);
            }
            {
                var a_Context = PdfDocument.Load(file);
                {
                    context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.HEADER, level, "[[[Info]]]");
                    {
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 1, "[[[File Name]]]", url);
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 1, "[[[File Size]]]", (new FileInfo(file)).Length.ToString());
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 1, "[[[Raw Format]]]", "PDF");
                    }
                }
                {
                    __Execute(context, level, a_Context, a_Name);
                }
                if (a_Context.PageCount > GetProperty(NAME.PROPERTY.PREVIEW_DOCUMENT_SIZE, true))
                {
                    context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.WARNING, level, "...");
                }
                {
                    context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOOTER, level, "[[[Pages]]]: " + a_Context.PageCount.ToString());
                    {
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, level + 1, "[[[Header]]]");
                        {
                            context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 2, "[[[Author]]]", a_Context.GetInformation()?.Author.ToString());
                            context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 2, "[[[Creator]]]", a_Context.GetInformation()?.Creator.ToString());
                            context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 2, "[[[Keywords]]]", a_Context.GetInformation()?.Keywords.ToString());
                            context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 2, "[[[Producer]]]", a_Context.GetInformation()?.Producer.ToString());
                            context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 2, "[[[Subject]]]", a_Context.GetInformation()?.Subject.ToString());
                            context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level + 2, "[[[Title]]]", a_Context.GetInformation()?.Title.ToString());
                        }
                    }
                    if (a_Context.Bookmarks?.Count > 0)
                    {
                        context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, level + 1, "[[[Bookmarks]]]");
                        {
                            __Execute(context, level + 2, a_Context.Bookmarks);
                        }
                    }
                }
                {
                    a_Context.Dispose();
                    GC.Collect();
                }
            }
            {
                context.
                    SetAlignment(NAME.ALIGNMENT.TOP).
                    SetFontState(NAME.FONT_STATE.NONE).
                    SetProgress(100).
                    SetUrlPreview(a_Name).
                    SendPreview(NAME.TYPE.INFO, url);
            }
        }

        private static int __GetSizeY(PdfDocument data)
        {
            var a_Result = 0;
            var a_Size1 = GetProperty(NAME.PROPERTY.PREVIEW_WIDTH, true);
            var a_Size2 = GetProperty(NAME.PROPERTY.PREVIEW_DOCUMENT_SIZE, true);
            for (var i = 0; (i < data.PageCount) && (i <= a_Size2); i++)
            {
                var a_Context = data.Render(i, 24, 24, true);
                if ((a_Context.Width > 0) && (a_Context.Height > 0))
                {
                    if (i == a_Size2)
                    {
                        a_Result += (CONSTANT.OUTPUT_PREVIEW_PAGE_BREAK * a_Context.Height) / 100;
                    }
                    else
                    {
                        a_Result += ((a_Size1 * a_Context.Height) / a_Context.Width) + CONSTANT.OUTPUT_PREVIEW_PAGE_INDENT;
                    }
                }
                {
                    a_Context.Dispose();
                }
            }
            return Math.Max(a_Result - CONSTANT.OUTPUT_PREVIEW_PAGE_INDENT, 0);
        }

        private static void __Execute(atom.Trace context, int level, PdfDocument data, string urlProxy)
        {
            var a_Size1 = GetProperty(NAME.PROPERTY.PREVIEW_WIDTH, true);
            var a_Size2 = GetProperty(NAME.PROPERTY.PREVIEW_DOCUMENT_SIZE, true);
            var a_Context = new Bitmap(a_Size1, __GetSizeY(data), PixelFormat.Format32bppArgb);
            {
                var a_Context1 = Graphics.FromImage(a_Context);
                var a_Index = 0;
                {
                    a_Context1.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    a_Context1.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    a_Context1.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    a_Context1.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.None;
                }
                for (var i = 0; (i < data.PageCount) && (i <= a_Size2); i++)
                {
                    var a_Context2 = data.Render(i, 96, 96, true);
                    if ((a_Context2.Width > 0) && (a_Context2.Height > 0))
                    {
                        var a_Size = new Rectangle(0, a_Index, a_Size1, ((a_Size1 * a_Context2.Height) / a_Context2.Width));
                        {
                            a_Context1.FillRectangle(new SolidBrush(Color.White), a_Size);
                            a_Context1.DrawImage(a_Context2, a_Size);
                            a_Context1.DrawRectangle(new Pen(Color.Black, 1), a_Size);
                        }
                        {
                            a_Index += a_Size.Height + CONSTANT.OUTPUT_PREVIEW_PAGE_INDENT;
                        }
                    }
                    {
                        a_Context2.Dispose();
                    }
                }
            }
            {
                a_Context.Save(urlProxy, ImageFormat.Png);
            }
            {
                var a_Size = (a_Context.Height + CONSTANT.OUTPUT_PREVIEW_ITEM_HEIGHT) / (CONSTANT.OUTPUT_PREVIEW_ITEM_HEIGHT + 1);
                {
                    a_Size = Math.Max(a_Size, CONSTANT.OUTPUT_PREVIEW_MIN_SIZE);
                }
                for (var i = 0; i < a_Size; i++)
                {
                    context.Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PREVIEW, level);
                }
            }
        }

        private static void __Execute(atom.Trace context, int level, PdfBookmarkCollection data)
        {
            if (data != null)
            {
                foreach (var a_Context in data)
                {
                    context.
                        SetComment("[[[Page]]]: " + (a_Context.PageIndex + 1).ToString(), "[[[Page index]]]").
                        Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level, a_Context.Title);
                    if (a_Context.Children?.Count > 0)
                    {
                        __Execute(context, level + 1, a_Context.Children);
                    }
                }
            }
        }
    };
}
