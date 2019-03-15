using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using iTextSharp.text;
using Org.BouncyCastle.Crypto.Digests;
using iTextSharp.text.pdf.security;
using com.itextpdf.text.pdf.security;
using Org.BouncyCastle.X509;
using System.IO;
using iTextSharp.text.pdf;

namespace PDFTools_iTextSharpLib
{
    public class iTextSharpUtil
    {
        // http://stackoverflow.com/questions/13618847/itextsharp-combining-two-pdf-documents-into-one

        public string ConcatFiles(List<string> files, string targetPath)
        {
            try
            {
                using (var ms = new System.IO.FileStream(GetFullPath(targetPath), System.IO.FileMode.Create))
                {
                    using (Document document = new Document())
                    {
                        using (PdfCopy copy = new PdfCopy(document, ms))
                        {
                            document.Open();
                            int n;
                            foreach (string sourceFile in files)
                            {
                                //PdfReader.unethicalreading = true; // http://stackoverflow.com/questions/17666577/opening-password-protected-pdf-file-with-itextsharp
                                //byte[] password = System.Text.ASCIIEncoding.ASCII.GetBytes("Secretinfo");
                                iTextSharp.text.pdf.PdfReader reader2 = new iTextSharp.text.pdf.PdfReader(GetFullPath(sourceFile)/*, password*/);
                                n = reader2.NumberOfPages;
                                for (int page = 0; page < n; )
                                {
                                    copy.AddPage(copy.GetImportedPage(reader2, ++page));
                                }
                                // close doc
                                reader2.Close();
                            }

                        }
                    }
                    return "";
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        // based on example http://itextsharp.sourceforge.net/examples/Concat.cs
        public string ConcatFilesOld(List<string> files, string targetPath)
        {
            try
            {
                if (files.Count > 0)
                {
                    string file = files[0];
                    iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(GetFullPath(file));
                    int n = reader.NumberOfPages;
                    iTextSharp.text.Document document = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(1));
                    iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new System.IO.FileStream(targetPath, System.IO.FileMode.Create));

                    reader.Close();
                    document.Open();
                    iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
                    iTextSharp.text.pdf.PdfImportedPage page;
                    int rotation;
                    foreach (string sourceFile in files)
                    {
                        int i = 0;
                        iTextSharp.text.pdf.PdfReader reader2 = new iTextSharp.text.pdf.PdfReader(GetFullPath(sourceFile));
                        n = reader2.NumberOfPages;
                        while (i < n)
                        {
                            i++;
                            document.SetPageSize(reader2.GetPageSizeWithRotation(i));
                            document.NewPage();
                            page = writer.GetImportedPage(reader2, i);
                            rotation = reader2.GetPageRotation(i);
                            if (rotation == 90)
                            {
                                cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader2.GetPageSizeWithRotation(i).Height);
                            }
                            else if ((rotation == 270))
                            {
                                cb.AddTemplate(page, 0f, 1f, -1f, 0f, reader2.GetPageSizeWithRotation(i).Width, 0f);
                            }
                            else
                            {
                                cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                            }

                        }
                        
                        writer.FreeReader(reader2);
                        reader2.Close();
                    }
                    
                    if (document.IsOpen())
                    {
                        document.CloseDocument();
                        document.Close();
                    }
                    return "";
                }
                else
                {
                    return "No files to process, use AddFile method";
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }

        public string AddText(string PathSource, string PathTarget, int x, int y, int selectedPage, String text = "text", int fontSize = 12)
        {
            return AddText(PathSource, PathTarget, x, y, selectedPage, text, 0, 0, 0, 0, fontSize, 1.0f);
        }

        public string AddText(string PathSource, string PathTarget, int x, int y, int selectedPage, String text, int angle, int red, int green, int blue, int fontSize, float opacity)
        {
            try
            {
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(PathSource);
                int n = reader.NumberOfPages;

                if (!(selectedPage > 0 && selectedPage <= n))
                {
                    return String.Format("Invalid Page {0}, the PDF has {1} pages", selectedPage, n);
                }

                iTextSharp.text.Document document = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(1));
                iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new System.IO.FileStream(PathTarget, System.IO.FileMode.Create));

                document.Open();
                iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
                iTextSharp.text.pdf.PdfImportedPage page;
                int rotation;
                int i = 0;
                // step 4: we add content
                while (i < n)
                {
                    i++;
                    document.NewPage();
                    page = writer.GetImportedPage(reader, i);
                    rotation = reader.GetPageRotation(i);
                    if (rotation == 90 || rotation == 270)
                    {
                        cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
                    }
                    else
                    {
                        cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                    }

                    if (i == selectedPage)
                        InsertText(cb, fontSize, x, y, text, angle, red, green, blue, opacity);
                }
                document.Close();
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        private void InsertText(PdfContentByte cb, int fontSize, int x, int y, String text, float angle, int red, int green, int blue, float opacity)
        {
            // Idea from https://forums.asp.net/t/1360567.aspx?iText+Sharp+and+text+position

            BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.BeginText();
            cb.SetFontAndSize(bf, fontSize);
            cb.SetRGBColorFill(red, green, blue);
            PdfGState gstate = new PdfGState();
            gstate.FillOpacity = opacity;
            cb.SetGState(gstate);
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, x, y, angle);
            cb.EndText();
        }

        public string AddSignature(string PathSource, string PathTarget, string CertPath, string CertPass, int lx = 100, int ly = 100, int ux = 250, int uy = 150, int page = 1, bool Visible = true)
        {

            try
            {
                Org.BouncyCastle.Crypto.AsymmetricKeyParameter Akp = null;
                Org.BouncyCastle.X509.X509Certificate[] Chain = null;

                string alias = null;
                Org.BouncyCastle.Pkcs.Pkcs12Store pk12;


                pk12 = new Org.BouncyCastle.Pkcs.Pkcs12Store(new System.IO.FileStream(CertPath, System.IO.FileMode.Open, System.IO.FileAccess.Read), CertPass.ToCharArray());

                IEnumerable aliases = pk12.Aliases;
                foreach (string aliasTemp in aliases)
                {
                    alias = aliasTemp;
                    if (pk12.IsKeyEntry(alias))
                        break;
                }

                Akp = pk12.GetKey(alias).Key;
                Org.BouncyCastle.Pkcs.X509CertificateEntry[] ce = pk12.GetCertificateChain(alias);
                Chain = new Org.BouncyCastle.X509.X509Certificate[ce.Length];
                for (int k = 0; k < ce.Length; ++k)
                    Chain[k] = ce[k].Certificate;

                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(PathSource);
                iTextSharp.text.pdf.PdfStamper st = iTextSharp.text.pdf.PdfStamper.CreateSignature(reader, new System.IO.FileStream(PathTarget, System.IO.FileMode.Create, System.IO.FileAccess.Write), '\0', null, true);
                iTextSharp.text.pdf.PdfSignatureAppearance sap = st.SignatureAppearance;
				
                if (Visible == true)
                {
					page = (page < 1 || page > reader.NumberOfPages) ? 1 : page;
                    sap.SetVisibleSignature(new iTextSharp.text.Rectangle(lx, ly, ux, uy), page, null);
				}

                sap.CertificationLevel = iTextSharp.text.pdf.PdfSignatureAppearance.CERTIFIED_NO_CHANGES_ALLOWED;

                // digital signature - http://itextpdf.com/examples/iia.php?id=222 

                IExternalSignature es = new PrivateKeySignature(Akp, "SHA-256"); // "BC"
                MakeSignature.SignDetached(sap, es, new X509Certificate[] { pk12.GetCertificate(alias).Certificate }, null, null, null, 0, CryptoStandard.CMS);

                st.Close();
                return "";
            }
            catch (Exception e)
            {
                return e.Message;

            }
        }

        public string AddSignature(string PathSource, string PathTarget, string CertPath, string CertPass, bool Visible)
		{
			return AddSignature(PathSource, PathTarget, CertPath, CertPass,100,100,250,150,1,true);
		}

		// based on http://itextsharp.sourceforge.net/examples/Encrypt.cs
		public string ModifyPermissions(string PathSource, string PathTarget, string UserPassword, List<int> Permissons)
        {
            try
            {
                int PDFpermisson = 0;
                foreach (int permisson in Permissons)
                {
                    PDFpermisson = PDFpermisson | permisson;
                }
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(PathSource);
                int n = reader.NumberOfPages;
                iTextSharp.text.Document document = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(1));
                iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new System.IO.FileStream(PathTarget, System.IO.FileMode.Create));
                writer.SetEncryption(iTextSharp.text.pdf.PdfWriter.STRENGTH128BITS, UserPassword, null, (int)PDFpermisson);
                // step 3: we open the document
                document.Open();
                iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
                iTextSharp.text.pdf.PdfImportedPage page;
                int rotation;
                int i = 0;
                // step 4: we add content
                while (i < n)
                {
                    i++;
                    document.SetPageSize(reader.GetPageSizeWithRotation(i));
                    document.NewPage();
                    page = writer.GetImportedPage(reader, i);
                    rotation = reader.GetPageRotation(i);
                    if (rotation == 90 || rotation == 270)
                    {
                        cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
                    }
                    else
                    {
                        cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                    }
                }
                document.Close();
                return "";

            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        public GeneXus.Utils.GXProperties GetFields(string PathSource)
        {
            GeneXus.Utils.GXProperties allFields = null;

            try
            {
                iTextSharp.text.pdf.PdfReader pdfReader = new iTextSharp.text.pdf.PdfReader(PathSource);

                allFields = new GeneXus.Utils.GXProperties();

                foreach (KeyValuePair<string,iTextSharp.text.pdf.AcroFields.Item> de in pdfReader.AcroFields.Fields)
                {
                    allFields.Add(de.Key.ToString(), de.Value.ToString());
                }
            }
            catch (Exception)
            {
                // Do nothing
            }
            return allFields;
        }

        // samples taken from 
        // http://www.c-sharpcorner.com/uploadfile/scottlysle/pdfgenerator_cs06162007023347am/pdfgenerator_cs.aspx
        // http://blog.codecentric.de/en/2010/08/pdf-generation-with-itext/
        public string SetFields(string PathSource, string PathTarget, System.Object myFields)
        {
            try
            {
                GeneXus.Utils.GXProperties Fields = (GeneXus.Utils.GXProperties)myFields;
                // create a new PDF reader based on the PDF template document
                iTextSharp.text.pdf.PdfReader pdfReader = new iTextSharp.text.pdf.PdfReader(PathSource);

                iTextSharp.text.pdf.PdfStamper pdfStamper = new iTextSharp.text.pdf.PdfStamper(pdfReader, new System.IO.FileStream(PathTarget, System.IO.FileMode.Create));

                GeneXus.Utils.GxKeyValuePair item = Fields.GetFirst();
                while (item != null)
                {
                    pdfStamper.AcroFields.SetField(item.Key, item.Value);
                    item = Fields.GetNext();
                }

                // flatten the form to remove editting options, set it to false to leave the form open to subsequent manual edits
                pdfStamper.FormFlattening = false;

                // close the pdf
                pdfStamper.Close();
                pdfReader.Close();
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }

        public int NumberOfPages(String PathSource)
        {
            int res = 0;
            try
            {
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(PathSource);
                res = reader.NumberOfPages;
            }
            catch (Exception)
            {
                res = -1; // error
            }
            return res;
        }

        public String TiffAsPDF(List<string> files, String targetPath)
        {
            try
            {
                if (files.Count > 0)
                {
                    //ByteArrayOutputStream outfile = new ByteArrayOutputStream();

                    iTextSharp.text.Document document = null;
                    iTextSharp.text.pdf.PdfWriter writer = null;

                    foreach (string str2 in files) // iterate over files
                    {

                        iTextSharp.text.Image image = null;
                        iTextSharp.text.pdf.RandomAccessFileOrArray ra = new iTextSharp.text.pdf.RandomAccessFileOrArray(GetFullPath(str2));
                        int pages = iTextSharp.text.pdf.codec.TiffImage.GetNumberOfPages(ra);

                        for (int iPage = 1; iPage <= pages; iPage++) // iterate over tiff pages
                        {

                            image = iTextSharp.text.pdf.codec.TiffImage.GetTiffImage(ra, iPage);

                            iTextSharp.text.Rectangle pageSize = new iTextSharp.text.Rectangle(image.PlainWidth, image.PlainHeight);
                            if (document == null)
                            {
                                document = new iTextSharp.text.Document(pageSize); // initialize with a PageSize
                                document.AddCreationDate();

                                writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new System.IO.FileStream(GetFullPath(targetPath), System.IO.FileMode.Create));
                                writer.StrictImageSequence = true;
                                document.Open();
                            }
                            else
                            {
                                document.SetPageSize(pageSize);
                            }
                            //tiff.scaleToFit(800, 600);  
                            document.Add(image);
                            document.NewPage();
                        }
                    }
                    document.Close();
                    //outfile.Flush();
                    return "";
                }
                return "No Tiff files to process";
            }
            catch (Exception exception)
            {
                return exception.Message;
            }
            finally
            {
                // Do nothing;
            }
        }

        public string SetSize(string sourceFile,string pSize, string targetPath)
        {
            try
            {
                string file = GetFullPath(sourceFile);
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(file);
                int n = reader.NumberOfPages;

                Rectangle pageSize = PageSize.GetRectangle(pSize);

                iTextSharp.text.Document document = new iTextSharp.text.Document(pageSize);
                iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, new System.IO.FileStream(GetFullPath(targetPath), System.IO.FileMode.Create));
                document.Open();
                iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
                iTextSharp.text.pdf.PdfImportedPage page;
                int rotation;
                int i = 0;

                reader = new iTextSharp.text.pdf.PdfReader(file);
                n = reader.NumberOfPages;
                while (i < n)
                {
                    i++;
                    document.SetPageSize(pageSize);
                    document.NewPage();
                    page = writer.GetImportedPage(reader, i);
                    rotation = reader.GetPageRotation(i);
                    if (rotation == 90)
                    {
                        cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height);
                    }
                    else if ((rotation == 270))
                    {
                        cb.AddTemplate(page, 0f, 1f, -1f, 0f, reader.GetPageSizeWithRotation(i).Width, 0f);
                    }
                    else
                    {
                        cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                    }

                }
                document.Close();
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }

        }

        public string GetFullPath(string f)
        {
            string fileFullPath = (System.IO.Path.IsPathRooted(f) ? f : Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, f));
            return fileFullPath;
        }

    }
}
