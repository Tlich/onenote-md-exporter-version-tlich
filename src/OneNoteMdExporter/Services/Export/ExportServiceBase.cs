using alxnbl.OneNoteMdExporter.Helpers;
using alxnbl.OneNoteMdExporter.Infrastructure;
using alxnbl.OneNoteMdExporter.Models;
using Microsoft.Office.Interop.OneNote;
using Serilog;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace alxnbl.OneNoteMdExporter.Services.Export
{
    public abstract class ExportServiceBase : IExportService
    {
        protected readonly AppSettings _appSettings;
        protected readonly Application _oneNoteApp;
        protected readonly ConverterService _convertServer;

        protected string _exportFormatCode;

        public ExportServiceBase(AppSettings appSettings, Application oneNoteApp, ConverterService converterService)
        {
            _appSettings = appSettings;
            _oneNoteApp = oneNoteApp;
            _convertServer = converterService;
        }


        protected string GetNotebookFolderPath(Notebook notebook)
            => Path.Combine(notebook.ExportFolder, notebook.GetNotebookPath());

        /// <summary>
        /// Return location in the export folder of an attachment file
        /// </summary>
        /// <param name="page"></param>
        /// <param name="attachId">Id of the attachment</param>
        /// <param name="oneNoteFilePath">Original filepath of the file in OneNote</param>
        /// <returns></returns>
        protected abstract string GetAttachmentFilePath(Attachement attachement);

        protected abstract string GetAttachmentFilePathOnPage(Attachement attachement);

        /// <summary>
        /// Get the md reference to the attachment
        /// </summary>
        /// <param name="attachement"></param>
        /// <returns></returns>
        protected abstract string GetAttachmentMdReference(Attachement attachement);

        protected abstract string GetResourceFolderPath(Node node);

        protected abstract string GetPageMdFilePath(Page page);


        public void ExportNotebook(Notebook notebook, string sectionNameFilter = "", string pageNameFilter = "")
        {
            //创建总文件夹
            notebook.ExportFolder = $"{Localizer.GetString("ExportFolder")}\\{_exportFormatCode}\\{notebook.GetNotebookPath()}-{DateTime.Now.ToString("yyyyMMdd HH-mm")}";
            CleanUpFolder(notebook); //创建笔记本名字的文件夹
            Directory.CreateDirectory(GetResourceFolderPath(notebook)); // 创建resource文件夹
            // Initialize hierarchy of the notebook from OneNote APIs:从OneNote api中初始化笔记本的层次结构
            try
            {
                _oneNoteApp.FillNodebookTree(notebook);
            }
            catch (Exception ex)
            {
                Log.Error(ex, Localizer.GetString("ErrorDuringNotebookProcessingNbTree"), notebook.Title, notebook.Id, ex.Message);
                return;
            }

            ExportNotebookInTargetFormat(notebook, sectionNameFilter, pageNameFilter); 
        }

        public abstract void ExportNotebookInTargetFormat(Notebook notebook, string sectionNameFilter = "", string pageNameFilter = "");

        private void CleanUpFolder(Notebook notebook)
        {
            // Cleanup Notebook export folder
            DirectoryHelper.ClearFolder(GetNotebookFolderPath(notebook));

            // Cleanup temp folder
            DirectoryHelper.ClearFolder(GetTmpFolder(notebook));
        }

        protected abstract void PreparePageExport(Page page);

        protected string GetTmpFolder(Node node)
        {
            return _appSettings.UserTempFolder ?
                Path.Combine(Path.GetTempPath(), node.GetNotebookPath()) :
                Path.Combine("tmp", node.GetNotebookPath());
        }

        /// <summary>
        /// Export a Page and its attachments
        /// </summary>
        /// <param name="page"></param>
        /// <returns></returns>
        protected void ExportPage(Page page) //导出页面
        {
            // Suffix page title：后缀页面标题
            EnsurePageUniquenessPerSection(page);

            var docxFileTmpFile = Path.Combine(GetTmpFolder(page), page.Id + ".docx"); 

            try
            {
                if (File.Exists(docxFileTmpFile))
                    File.Delete(docxFileTmpFile);

                PreparePageExport(page); //创建页面文件夹

                Log.Debug($"{page.OneNoteId}: start OneNote docx publish");

                // Request OneNote to export the page into a DocX file：请求OneNote将页面导出为DocX文件
                _oneNoteApp.Publish(page.OneNoteId, Path.GetFullPath(docxFileTmpFile), PublishFormat.pfWord);

                Log.Debug($"{page.OneNoteId}: success");

                // Convert docx file into Md using PanDoc：利用PanDoc.exe将Word转为Md；pageMd就是word转为md的变量，但是不包括图片
                var pageMd = _convertServer.ConvertDocxToMd(page, docxFileTmpFile, GetTmpFolder(page)); 

                 
                 /*if (_appSettings.Debug)
                 {
                    // If debug mode enabled, copy the page docx file next to the page md file
                    // 如果启用了调试模式，保存页面的word文件
                    var docxFilePath = Path.ChangeExtension(GetPageMdFilePath(page), "docx");
                    File.Copy(docxFileTmpFile, docxFilePath);
                 }*/
                

                File.Delete(docxFileTmpFile); //删除临时docx

                // Copy images extracted from DocX to Export folder and add them in the list of attachments of the page
                // 将从Word中提取的图像复制到Export文件夹，并将它们添加到页面的附件列表中
                try
                {
                    //保存图片到resource文件夹
                    ExtractImagesToResourceFolder(page, ref pageMd, _appSettings.PostProcessingMdImgRef);
                }
                catch (COMException ex)
                {
                    if (ex.Message.Contains("0x800706BE"))
                    {
                        LogError(page, ex, Localizer.GetString("ErrorWhileStartingOnenote"));
                    }
                    else
                        LogError(page, ex, String.Format(Localizer.GetString("ErrorDuringOneNoteExport"), ex.Message));
                }
                catch (Exception ex)
                {
                    LogError(page, ex, Localizer.GetString("ErrorImageExtract"));
                }

                // Export all file attachments and get updated page markdown including md reference to attachments
                // 导出所有的文件附件，并得到更新的页面markdown，包括对附件的md引用
                ExportPageAttachments(page, ref pageMd);

                // Apply post processing to Page Md content：对Page Md内容应用后处理
                _convertServer.PageMdPostConvertion(page, ref pageMd);

                WritePageMdFile(page, pageMd);
            }
            catch (Exception ex)
            {
                LogError(page, ex, String.Format(Localizer.GetString("ErrorDuringPageProcessing"), page.TitleWithPageLevelTabulation, page.Id, ex.Message));
            }
        }

        private void LogError(Page p, Exception ex, string message)
        {
            Log.Warning($"Page '{p.GetPageFileRelativePath()}': {message}");
            //Log.Debug(ex, ex.Message);
            Log.Information(ex, ex.Message);
        }

        /// <summary>
        /// Final class needs to implement logic to write the md file of the page in the export folder
        /// </summary>
        /// <param name="page">The page</param>
        /// <param name="pageMd">Markdown content of the page</param>
        protected abstract void WritePageMdFile(Page page, string pageMd);


        /// <summary>
        /// Create attachment files in export folder, and update page's markdown to insert md reference that link to the attachment files
        /// </summary>
        /// <param name="page"></param>
        /// <param name="pageMdFileContent">Markdown content of the page</param>
        private void ExportPageAttachments(Page page, ref string pageMdFileContent)
        {
            foreach (Attachement attach in page.Attachements)
            {
                if (attach.Type == AttachementType.File)
                {
                    EnsureAttachmentFileIsNotUsed(page, attach);

                    var exportFilePath = GetAttachmentFilePath(attach);

                    // Copy attachment file into export folder
                    File.Copy(attach.ActualSourceFilePath, exportFilePath);
                    //File.SetAttributes(exportFilePath, FileAttributes.Normal); // Prevent exception during deletation of export directory

                    // Update page markdown to insert md references to attachments
                    InsertPageMdAttachmentReference(ref pageMdFileContent, attach, GetAttachmentMdReference);
                }

                FinalizeExportPageAttachemnts(page, attach);
            }
        }

        
        /// <summary>
        /// Final class needs to implement logic to write the md file of the attachment file in the export folder (if needed)
        /// </summary>
        /// <param name="page">The page</param>
        /// <param name="attachment">The attachment</param>
        protected abstract void FinalizeExportPageAttachemnts(Page page, Attachement attachment);


        /// <summary>
        /// Replace the tag <<FileName>> generated by OneNote by a markdown link referencing the attachment
        /// </summary>
        /// <param name="pageMdFileContent"></param>
        /// <param name="attach"></param>
        private static void InsertPageMdAttachmentReference(ref string pageMdFileContent, Attachement attach, Func<Attachement, string> getAttachMdReferenceMethod)
        {
            var pageMdFileContentModified = Regex.Replace(pageMdFileContent, "(\\\\<){2}(?<fileName>.*)(>\\\\>)", delegate (Match match)
            {
                var refFileName = match.Groups["fileName"]?.Value ?? "";
                var attachOriginalFileName = attach.OneNotePreferredFileName ;
                var attachMdRef = getAttachMdReferenceMethod(attach);

                if (refFileName.Equals(attachOriginalFileName))
                {
                    // reference found is corresponding to the attachment being processed
                    return $"[{attachOriginalFileName}]({attachMdRef})";
                }
                else
                {
                    // not the current attachmeent, ignore
                    return match.Value;
                }
            });

            pageMdFileContent = pageMdFileContentModified;
        }


        /// <summary>
        /// Replace PanDoc IMG HTML tag by markdown reference and copy image file into notebook export directory
        /// </summary>
        /// <param name="page">Section page</param>
        /// <param name="mdFileContent">Contennt of the MD file</param>
        /// <param name="resourceFolderPath">The path to the notebook folder where store attachments</param>
        /// <param name="postProcessingMdImgRef">If false, markdown reference to image will not be inserted</param>
        /// <param name="getImgMdReferenceMethod">The method that returns the md reference of an image attachment</param>
        public void ExtractImagesToResourceFolder(Page page, ref string mdFileContent, bool postProcessingMdImgRef)
        {
            // Replace <IMG> tags by markdown references：用markdown引用替换<IMG>标签，这估计是一个循环
            // delegate (Match match)是改名的核心
            var pageTxtModified = Regex.Replace(mdFileContent, "<img [^>]+/>", delegate (Match match)
            {

                string imageTag = match.ToString();

                // http://regexstorm.net/tester
                string regexImgAttributes = "<img src=\"(?<src>[^\"]+)\".* />";

                MatchCollection matchs = Regex.Matches(imageTag, regexImgAttributes, RegexOptions.IgnoreCase);
                Match imgMatch = matchs[0]; //matchs是提取的图片内容，matchs[0]是图片本体的内容

                var panDocHtmlImgTagPath = Path.GetFullPath(imgMatch.Groups["src"].Value); 

                Attachement imgAttach = page.ImageAttachements.Where(img => PathExtensions.PathEquals(img.ActualSourceFilePath, panDocHtmlImgTagPath)).FirstOrDefault();

                // Only add a new attachment if this is the first time the image is referenced in the page
                // 仅当第一次在页面中引用图像时，才添加新附件
                if (imgAttach == null)
                {
                    // Add a new attachmeent to current page
                    imgAttach = new Attachement(page)
                    {
                        Type = AttachementType.Image,
                    };

                    imgAttach.ActualSourceFilePath = Path.GetFullPath(panDocHtmlImgTagPath);
                    // Not really a use file path but a PanDoc temp file：不是真正的使用文件路径，而是一个PanDoc临时文件
                    imgAttach.OriginalUserFilePath = Path.GetFullPath(panDocHtmlImgTagPath);

                    page.Attachements.Add(imgAttach);

                    EnsureAttachmentFileIsNotUsed(page, imgAttach); //这里出现了OverrideExportFilePath
                }

                var attachRef = GetAttachmentMdReference(imgAttach);
                var refLabel = Path.GetFileNameWithoutExtension(imgAttach.ActualSourceFilePath);

                return $"![{refLabel}]({attachRef})";
            });


            // Move attachements file into output ressource folder and delete tmp file：将附件文件移动到输出资源文件夹并删除tmp文件
            // In case of dupplicate files, suffix attachment file name：如有重复文件，请以附件文件名作为后缀
            foreach (var attach in page.ImageAttachements) //attach就是里面的图片，page就是
            {
                //GetAttachmentFilePath(attach) 获取attach.ActualSourceFilePath，即为图片路径
                File.Copy(attach.ActualSourceFilePath, GetAttachmentFilePathOnPage(attach));
                File.Delete(attach.ActualSourceFilePath);
            }


            if (postProcessingMdImgRef)
            {
                mdFileContent = pageTxtModified;
            }
        }

        /// <summary>
        /// Suffix the attachment file name if it conflicits with an other attachement previously attached to the notebook export
        /// </summary>
        /// <param name="page">The parent Page</param>
        /// <param name="attach">The attachment</param>
        private void EnsureAttachmentFileIsNotUsed(Page page, Attachement attach)
        {
            var notUseFileNameFound = false;
            var cmpt = 0;
            var attachmentFilePath = GetAttachmentFilePath(attach);

            while (!notUseFileNameFound)
            {
                var candidateFilePath = cmpt == 0 ? attachmentFilePath :
                    $"{Path.ChangeExtension(attachmentFilePath, null)}-{cmpt}{Path.GetExtension(attachmentFilePath)}";

                var attachmentFileNameAlreadyUsed = page.GetNotebook().GetAllAttachments().Any(a => a != attach && PathExtensions.PathEquals(GetAttachmentFilePath(a), candidateFilePath));
                
                if (!attachmentFileNameAlreadyUsed)
                {
                    if (cmpt > 0)
                        attach.OverrideExportFilePath = candidateFilePath;

                    notUseFileNameFound = true;
                }
                else
                    cmpt++;
            }

        }


        /// <summary>
        /// Suffix the page file name if it conflicits with an other page previously attached to the notebook export
        /// </summary>
        /// <param name="page">The parent Page</param>
        /// <param name="attach">The attachment</param>
        private void EnsurePageUniquenessPerSection(Page page)
        {
            var notUseFileNameFound = false;
            var cmpt = 0;
            var pageFilePath = GetPageMdFilePath(page);

            while (!notUseFileNameFound)
            {
                var candidateFilePath = cmpt == 0 ? pageFilePath :
                    $"{Path.ChangeExtension(pageFilePath, null)}-{cmpt}.md";

                var attachmentFileNameAlreadyUsed = page.Parent.Childs.OfType<Page>().Any(p => p != page && PathExtensions.PathEquals(GetPageMdFilePath(p), candidateFilePath));

                if (!attachmentFileNameAlreadyUsed)
                {
                    if (cmpt > 0)
                        page.OverridePageFilePath = candidateFilePath;

                    notUseFileNameFound = true;
                }
                else
                    cmpt++;
            }

        }

    }
}
