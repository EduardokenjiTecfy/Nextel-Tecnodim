using GdPicture;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace CastPdfTiffNextel
{
    class Program
    {
        public static void varios(String entrada)
        {
            var dialog = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Abrir Documentos",
                Filter = "Documentos Pdf|*.pdf"
            };
            using (dialog)
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {

                    foreach (var item in dialog.FileNames)
                    {
                        if (File.Exists(entrada + @"\" + Path.GetFileName(item)))
                        {
                            var filename = Path.GetFileName(item);
                            Console.WriteLine(filename);
                            File.Delete(entrada + @"\" + Path.GetFileName(item));
                        }


                        File.Move(item, entrada + @"\" + Path.GetFileName(item));
                    }

                }
            }
        }

        public static void moveFile(string pdf)

        {

            if (!File.Exists(ConfigurationManager.AppSettings["Destino"].ToString() + Path.GetFileName(pdf)))
            {
                var caminho = ConfigurationManager.AppSettings["Destino"].ToString();
                File.Move(pdf, ConfigurationManager.AppSettings["Destino"].ToString() + Path.GetFileName(pdf));
            }
        }
        private static string GetCurrentDirectory()
        {
            var executingAssemblyPath = new System.Uri(Assembly.GetExecutingAssembly().CodeBase).AbsolutePath;
            var directory = new DirectoryInfo(Path.GetDirectoryName(executingAssemblyPath)).FullName;

            return Uri.UnescapeDataString(directory);
        }
        public static Stream GerarDocumentoPesquisavel(GdPictureImaging _gdPictureImaging, GdPicturePDF _gdPicturePDF, string documento, bool pdfa = true, string idioma = "por", string titulo = null, string autor = null, string assunto = null, string palavrasChaves = null, string criador = null, int dpi = 200)
        {
            int pdfOcrID = 0;
            Stream documentoConvertido = new MemoryStream();

            GdPictureStatus status = _gdPicturePDF.LoadFromFile(documento, true);
            var _diretorioIdiomas = @"C:\tfs\bpo\TecnoDim\Branch\Main\Source\TecnoDIM\TecnoDim.Server.ActivityDesigner\bin\Debug\GdPicture\Idiomas";

            pdfOcrID = _gdPictureImaging.PdfOCRStart(@"C:\Users\Administrador\Desktop\a\testevtnc.pdf", pdfa, titulo, autor, assunto, palavrasChaves, criador);

            int totalPaginas = _gdPicturePDF.GetPageCount();
            for (int i = 1; i <= totalPaginas; i++)
            {
                _gdPicturePDF.SelectPage(i);
                int rasterPageID = _gdPicturePDF.RenderPageToGdPictureImageEx(dpi, true);
                if (rasterPageID != 0)
                {
                    _gdPictureImaging.PdfAddGdPictureImageToPdfOCR(pdfOcrID, rasterPageID, idioma, _diretorioIdiomas, "");

                    _gdPictureImaging.ReleaseGdPictureImage(rasterPageID);
                }
            }

            _gdPictureImaging.PdfOCRStop(pdfOcrID);
            _gdPicturePDF.CloseDocument();

            return documentoConvertido;
        }

        static void Main(string[] args)
        {


            int ImageID = 0;
            bool bContinue = false;
            int PdfID = 0;
            GdPictureImaging oGdPictureImaging = new GdPictureImaging();
            GdPicturePDF oGdPicturePDF = new GdPicturePDF();
            oGdPicturePDF.SetLicenseNumber("4118106456693265856441854");
            oGdPictureImaging.SetLicenseNumber("4118106456693265856441854");


            Console.WriteLine(args[0].ToString());









            Console.WriteLine("TecnoDim está processando os arquivos: ");
            var argumentos = args[0].ToString().Split(';');



            //    GdPictureImaging oGdPictureImaging = new GdPictureImaging();

            oGdPictureImaging.SetLicenseNumber("4118106456693265856441854");


            var entrada = argumentos[0];
            var saida = argumentos[1];
            var concessionaria = argumentos[2];
            string lote = "";
            string nomeUsuario = "";
            int dpi = Convert.ToInt32(argumentos[6]);
            string debito_nome = "";

            var arquivos = Directory.GetFiles(saida).ToList();
            foreach (var item in arquivos)
            {
                var arquivoExtensao = Path.GetExtension(item);
                if (arquivoExtensao == ".tif" || arquivoExtensao == ".tiff")
                {
                    File.Delete(item);
                }
            }



            PixelFormat pixelFormat = (PixelFormat)Enum.Parse(typeof(PixelFormat), argumentos[5], true);
            TiffCompression tiffCompression = (TiffCompression)Enum.Parse(typeof(TiffCompression), argumentos[7], true);
            string debito = argumentos[9];
            var GrayScale = Convert.ToInt32(argumentos[8]);



            if (debito == "Debito")
            {

                debito_nome = "1";
            }
            else
            {

                debito_nome = "0";
            }
            if (argumentos.Length > 3)
            {
                lote = argumentos[3];
                nomeUsuario = argumentos[4];



            }

            switch (concessionaria)
            {
                case "OUTROS":


                    var trd = new Thread(() => varios(entrada)
                );
                    int arquivo = 0;
                    trd.IsBackground = true;
                    trd.SetApartmentState(ApartmentState.STA);
                    trd.Start();
                    while (trd.IsAlive)
                    {

                    }



                    foreach (var pdf in Directory.GetFiles(entrada))
                    {

                        Console.WriteLine(Path.GetFileName(pdf));




                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int multipageHandle = 0;
                            int imageId = 0;
                            int imagem = 0;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {

                                imagem++;



                                oGdPicturePDF.SelectPage(imagem);
                                imageId = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                //oGdPictureImaging.ConvertBitonalToGrayScale(imageId, GrayScale);
                                oGdPictureImaging.ConvertTo1Bpp(imageId, (byte)130);




                                if (i == 0)
                                {

                                    multipageHandle = imageId;
                                    // oGdPictureImaging.TiffSaveAsMultiPageFile(multipageHandle, saida + @"\" + Path.GetFileName(pdf) + "_.tif", tiffCompression);
                                    oGdPictureImaging.TiffSaveAsMultiPageFile(multipageHandle, saida + @"\" + "_" + lote.ToString() + "_" + arquivo.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);
                                    arquivo++;
                                }
                                else //If it is second image or more, add it to previously created multipage tiff file
                                {
                                    oGdPictureImaging.TiffAddToMultiPageFile(multipageHandle, imageId);
                                    //Release current image to minimize memory usage
                                    oGdPictureImaging.ReleaseGdPictureImage(imageId);
                                }

                            }

                            oGdPictureImaging.TwainCloseSource(); //Closing the source
                            oGdPictureImaging.TiffCloseMultiPageFile(multipageHandle); //Closing multipage file
                            oGdPictureImaging.ReleaseGdPictureImage(multipageHandle); //Releasing the multipage image 
                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }
                    break;



                case "ENERGISA":

                    var tsrdenergisa = new Thread(() => varios(entrada)
                     );

                    tsrdenergisa.IsBackground = true;
                    tsrdenergisa.SetApartmentState(ApartmentState.STA);
                    tsrdenergisa.Start();
                    while (tsrdenergisa.IsAlive)
                    {

                    }


                    int arquivotsrdenergisa = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int multipageHandle = 0;
                            int imageId = 0;
                            int imagem = 0;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {

                                imagem++;


                                oGdPicturePDF.SelectPage(imagem);
                                imageId = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                oGdPictureImaging.ConvertBitonalToGrayScale(imageId, GrayScale);
                                oGdPictureImaging.ConvertTo1Bpp(imageId, 134);




                                if (i % 2 == 0)
                                {
                                    multipageHandle = imageId;
                                    oGdPictureImaging.TiffSaveAsMultiPageFile(multipageHandle, saida + @"\" + "_" + lote.ToString() + "_" + arquivotsrdenergisa.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);
                                    arquivotsrdenergisa++;
                                }
                                else //If it is second image or more, add it to previously created multipage tiff file
                                {
                                    oGdPictureImaging.TiffAddToMultiPageFile(multipageHandle, imageId);
                                    //Release current image to minimize memory usage
                                    oGdPictureImaging.ReleaseGdPictureImage(imageId);
                                    oGdPictureImaging.TwainCloseSource(); //Closing the source
                                    oGdPictureImaging.TiffCloseMultiPageFile(multipageHandle); //Closing multipage file
                                    oGdPictureImaging.ReleaseGdPictureImage(multipageHandle); //Releasing the multipage image 
                                }

                            }


                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }
                    break;
                case "CPFL":

                    var tsrdCPFL = new Thread(() => varios(entrada)
                     );

                    tsrdCPFL.IsBackground = true;
                    tsrdCPFL.SetApartmentState(ApartmentState.STA);
                    tsrdCPFL.Start();
                    while (tsrdCPFL.IsAlive)
                    {

                    }


                    int arquivotsrdCPFL = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int multipageHandle = 0;
                            int imageId = 0;
                            int imagem = 0;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {

                                imagem++;


                                oGdPicturePDF.SelectPage(imagem);
                                imageId = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                oGdPictureImaging.ConvertBitonalToGrayScale(imageId, GrayScale);
                                oGdPictureImaging.ConvertTo1Bpp(imageId, 196);




                                if (i == 0)
                                {
                                    multipageHandle = imageId;
                                    oGdPictureImaging.TiffSaveAsMultiPageFile(multipageHandle, saida + @"\" + "_" + lote.ToString() + "_" + arquivotsrdCPFL.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);
                                    arquivotsrdCPFL++;
                                }
                                else //If it is second image or more, add it to previously created multipage tiff file
                                {
                                    oGdPictureImaging.TiffAddToMultiPageFile(multipageHandle, imageId);
                                    //Release current image to minimize memory usage
                                    oGdPictureImaging.ReleaseGdPictureImage(imageId);
                                }

                            }

                            oGdPictureImaging.TwainCloseSource(); //Closing the source
                            oGdPictureImaging.TiffCloseMultiPageFile(multipageHandle); //Closing multipage file
                            oGdPictureImaging.ReleaseGdPictureImage(multipageHandle); //Releasing the multipage image 
                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }
                    break;


                case "COPEL":
                    var trdCOPEL = new Thread(() => varios(entrada)
                     );

                    trdCOPEL.IsBackground = true;
                    trdCOPEL.SetApartmentState(ApartmentState.STA);
                    trdCOPEL.Start();
                    while (trdCOPEL.IsAlive)
                    {

                    }


                    int contadorCOPEL = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int multipageHandle = 0;
                            int imageId = 0;
                            //loop through pages
                            for (int i = 5; i < oGdPicturePDF.GetPageCount(); i++)
                            {
                                oGdPictureImaging.ConvertTo1Bpp(imageId);
                                oGdPicturePDF.SelectPage(i);
                                imageId = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);

                                oGdPictureImaging.ConvertTo1Bpp(imageId, (byte)(157));
                                //Checking if image was loaded
                                if (imageId != 0)
                                {
                                    if (i % 2 != 0)
                                    {

                                        multipageHandle = imageId;

                                        oGdPictureImaging.TiffSaveAsMultiPageFile(multipageHandle, saida + @"\" + "_" + lote.ToString() + "_" + contadorCOPEL.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);
                                        contadorCOPEL++;
                                    }
                                    else //If it is second image or more, add it to previously created multipage tiff file
                                    {
                                        oGdPictureImaging.TiffAddToMultiPageFile(multipageHandle, imageId);
                                        oGdPictureImaging.ReleaseGdPictureImage(imageId);
                                        oGdPictureImaging.TiffCloseMultiPageFile(imageId);
                                        oGdPictureImaging.ReleaseGdPictureImage(imageId);
                                        oGdPictureImaging.ReleaseGdPictureImage(multipageHandle);
                                        oGdPictureImaging.TiffCloseMultiPageFile(multipageHandle);
                                        oGdPictureImaging.ReleaseGdPictureImage(multipageHandle);

                                    }
                                }

                            }


                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }
                    break;

                case "CEEE":
                    var trdCEEE = new Thread(() => varios(entrada)
                     );

                    trdCEEE.IsBackground = true;
                    trdCEEE.SetApartmentState(ApartmentState.STA);
                    trdCEEE.Start();
                    while (trdCEEE.IsAlive)
                    {

                    }


                    int arquivoCEEE = 0;

                    arquivoCEEE = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {

                        Console.WriteLine(Path.GetFileName(pdf));




                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int multipageHandle = 0;
                            int imageId = 0;
                            int imagem = 0;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {

                                imagem++;



                                oGdPicturePDF.SelectPage(imagem);
                                imageId = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                //oGdPictureImaging.ConvertBitonalToGrayScale(imageId, GrayScale);
                                oGdPictureImaging.ConvertTo1Bpp(imageId, (byte)182);




                                if (i == 0)
                                {
                                    multipageHandle = imageId;
                                    oGdPictureImaging.TiffSaveAsMultiPageFile(multipageHandle, saida + @"\" + "_" + lote.ToString() + "_" + arquivoCEEE.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);
                                    arquivoCEEE++;
                                }
                                else //If it is second image or more, add it to previously created multipage tiff file
                                {
                                    oGdPictureImaging.TiffAddToMultiPageFile(multipageHandle, imageId);
                                    //Release current image to minimize memory usage
                                    oGdPictureImaging.ReleaseGdPictureImage(imageId);
                                }

                            }

                            oGdPictureImaging.TwainCloseSource(); //Closing the source
                            oGdPictureImaging.TiffCloseMultiPageFile(multipageHandle); //Closing multipage file
                            oGdPictureImaging.ReleaseGdPictureImage(multipageHandle); //Releasing the multipage image 
                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }
                    break;


                case "LIGHT":
                    var trdLIGHT = new Thread(() => varios(entrada)
                     );

                    trdLIGHT.IsBackground = true;
                    trdLIGHT.SetApartmentState(ApartmentState.STA);
                    trdLIGHT.Start();
                    while (trdLIGHT.IsAlive)
                    {

                    }


                    int arquivotrdLIGHT = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        int multiTiffID = 0;

                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);

                        if (status == GdPictureStatus.OK)
                        {
                            int image = 1;

                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {



                                var dirsaida = saida + @"\" + "_" + lote.ToString() + "_" + arquivotrdLIGHT.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";
                                arquivotrdLIGHT++;
                                oGdPicturePDF.SelectPage(image);
                                image++;
                                //render selected page to GdPictureImage identifier
                                int rasterizedPageID = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                multiTiffID = rasterizedPageID;

                                oGdPictureImaging.ConvertBitonalToGrayScale(rasterizedPageID, GrayScale);
                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID, 175);
                                //make the image a new multipage tiff image
                                status = oGdPictureImaging.TiffSaveAsMultiPageFile(multiTiffID, dirsaida, tiffCompression);
                                //release old image identifier
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageID);

                                oGdPictureImaging.TiffCloseMultiPageFile(multiTiffID);

                                oGdPictureImaging.ReleaseGdPictureImage(multiTiffID);

                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID);
                                System.Drawing.Image myImagesflip = Image.FromFile(dirsaida);
                                if (myImagesflip.Height > myImagesflip.Width)
                                {

                                    myImagesflip.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                    myImagesflip.Save(dirsaida);


                                }

                                var s = MergeTwoImagesXY(myImagesflip, 0, 0, Image.FromFile(saida + @"\" + "ROTULO.png"));
                                myImagesflip.Dispose();
                                int cm = oGdPictureImaging.CreateGdPictureImageFromBitmap(s);
                                oGdPictureImaging.ConvertTo1Bpp(cm);
                                oGdPictureImaging.SaveAsTIFF(cm, dirsaida, tiffCompression);
                                oGdPictureImaging.ReleaseGdPictureImage(cm);
                                s.Dispose();

                            }


                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }

                    break;


                case "RGESUL":
                    var trdRGESUL = new Thread(() => varios(entrada)
                     );

                    trdRGESUL.IsBackground = true;
                    trdRGESUL.SetApartmentState(ApartmentState.STA);
                    trdRGESUL.Start();
                    while (trdRGESUL.IsAlive)
                    {

                    }


                    int arquivoRGESUL = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int multipageHandle = 0;
                            int imageId = 0;
                            int imagem = 0;
                            //loop through pages

                            string caminho = "";

                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {

                                imagem++;




                                oGdPicturePDF.SelectPage(imagem);
                                imageId = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                oGdPictureImaging.ConvertTo1Bpp(imageId, (byte)192);




                                if (i == 0)
                                {
                                    multipageHandle = imageId;
                                    caminho = saida + @"\" + "_" + lote.ToString() + "_" + i.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";
                                    oGdPictureImaging.TiffSaveAsMultiPageFile(multipageHandle, saida + @"\" + "_" + lote.ToString() + "_" + arquivoRGESUL.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);
                                    arquivoRGESUL++;
                                    GdPictureImaging oGdPictureImagingCorte = new GdPictureImaging();
                                    GdPicturePDF oGdPicturePDFCORTE = new GdPicturePDF();
                                    oGdPicturePDFCORTE.SetLicenseNumber("4118106456693265856441854");
                                    oGdPictureImagingCorte.SetLicenseNumber("4118106456693265856441854");

                                    int imageIdCorte = 0;


                                    oGdPicturePDFCORTE.SelectPage(imagem);
                                    imageIdCorte = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                    oGdPictureImagingCorte.ConvertTo1Bpp(imageIdCorte, (byte)140);
                                    oGdPictureImagingCorte.TiffSaveAsMultiPageFile(imageIdCorte, saida + @"\" + "SaidaCorte1.tif", tiffCompression);

                                    oGdPictureImagingCorte.TwainCloseSource(); //Closing the source
                                    oGdPictureImagingCorte.TiffCloseMultiPageFile(imageIdCorte); //Closing multipage file
                                    oGdPictureImagingCorte.ReleaseGdPictureImage(imageIdCorte); //Releasing the multipage image 
                                    oGdPicturePDFCORTE.CloseDocument();



                                }
                                else //If it is second image or more, add it to previously created multipage tiff file
                                {
                                    oGdPictureImaging.TiffAddToMultiPageFile(multipageHandle, imageId);
                                    //Release current image to minimize memory usage
                                    oGdPictureImaging.ReleaseGdPictureImage(imageId);
                                }





                            }
                            oGdPictureImaging.ReleaseGdPictureImage(imageId);
                            oGdPictureImaging.TwainCloseSource(); //Closing the source
                            oGdPictureImaging.TiffCloseMultiPageFile(multipageHandle); //Closing multipage file
                            oGdPictureImaging.ReleaseGdPictureImage(multipageHandle); //Releasing the multipage image 
                            oGdPicturePDF.CloseDocument();
                            System.Drawing.Image myImage = Image.FromFile(saida + @"\saidaCorte1.tif");


                            int f = myImage.Width;
                            Bitmap croppedBitmap = new Bitmap(myImage);
                            croppedBitmap.SetResolution(dpi, dpi);
                            croppedBitmap = croppedBitmap.Clone(
                                        new Rectangle(0, 0, f, 697),
                                        pixelFormat);
                            int wd = myImage.Width - croppedBitmap.Width, hd = myImage.Height - croppedBitmap.Height;

                            int imgteste = oGdPictureImaging.CreateGdPictureImageFromBitmap(croppedBitmap);

                            oGdPictureImaging.ConvertTo1Bpp(imgteste);

                            oGdPictureImaging.SaveAsTIFF(imgteste, saida + @"\CROP1.tif", tiffCompression);
                            oGdPictureImaging.ReleaseGdPictureImage(imgteste);
                            myImage.Dispose();
                            croppedBitmap.Dispose();


                            Image myImageOutroCorte = Image.FromFile(saida + @"\" + "_" + lote.ToString() + "_" + (arquivoRGESUL - 1).ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif");
                            Bitmap croppedBitmapCorte2 = new Bitmap(myImageOutroCorte);
                            croppedBitmapCorte2.SetResolution(dpi, dpi);
                            var cropp2 = croppedBitmapCorte2.Clone(new Rectangle(0, 697, f, hd), pixelFormat);

                            int imgtesteCorte2 = oGdPictureImaging.CreateGdPictureImageFromBitmap(cropp2);
                            cropp2.Dispose();
                            croppedBitmapCorte2.Dispose();
                            oGdPictureImaging.ConvertTo1Bpp(imgtesteCorte2);

                            oGdPictureImaging.SaveAsTIFF(imgtesteCorte2, saida + @"\CROP2.tif", tiffCompression);
                            oGdPictureImaging.ReleaseGdPictureImage(imgtesteCorte2);


                            myImageOutroCorte.Dispose();
                            croppedBitmap.Dispose();
                            var d = MergeTwoImages(Image.FromFile(saida + @"\" + "CROP1.tif"), Image.FromFile(saida + @"\" + "CROP2.tif"));

                            d.Save(saida + @"\" + "_" + lote.ToString() + "_" + (arquivoRGESUL - 1).ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif" + "f");
                            d.Dispose();
                            int join = oGdPictureImaging.TiffCreateMultiPageFromFile(caminho);
                            oGdPictureImaging.ConvertTo1Bpp(join);
                            oGdPictureImaging.SaveAsTIFF(join, saida + @"\" + "_" + lote.ToString() + "_" + (arquivoRGESUL - 1) + "_" + nomeUsuario + "_" + debito_nome + "_.tif", TiffCompression.TiffCompressionCCITT4);
                            oGdPictureImaging.ReleaseGdPictureImage(join);

                            File.Delete(saida + @"\" + "out.tif");
                            File.Delete(saida + @"\" + "CROP2.tif");
                            File.Delete(saida + @"\" + "CROP1.tif");
                            File.Delete(saida + @"\" + "saidaCorte1.tif");


                        }

                        moveFile(pdf);
                    }

                    foreach (var item in Directory.GetFiles(saida))
                    {
                        if (Path.GetExtension(item) == ".tif")
                        {
                            File.Delete(item);
                        }
                    }


                    break;


                case "FATURAO":
                    var trdFATURAO = new Thread(() => varios(entrada)
                     );

                    trdFATURAO.IsBackground = true;
                    trdFATURAO.SetApartmentState(ApartmentState.STA);
                    trdFATURAO.Start();
                    while (trdFATURAO.IsAlive)
                    {

                    }


                    int arquivoFATURAO = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        int multiTiffID = 0;
                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int image = 1;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {



                                var dirsaida = saida + @"\" + "_" + lote.ToString() + "_" + arquivoFATURAO.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";
                                arquivoFATURAO++;
                                oGdPicturePDF.SelectPage(image);
                                image++;
                                //render selected page to GdPictureImage identifier
                                int rasterizedPageID = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                multiTiffID = rasterizedPageID;

                                oGdPictureImaging.ConvertBitonalToGrayScale(rasterizedPageID, GrayScale);
                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID, 175);
                                //make the image a new multipage tiff image
                                status = oGdPictureImaging.TiffSaveAsMultiPageFile(multiTiffID, dirsaida, tiffCompression);
                                //release old image identifier
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageID);

                                oGdPictureImaging.TiffCloseMultiPageFile(multiTiffID);

                                oGdPictureImaging.ReleaseGdPictureImage(multiTiffID);

                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID);

                            }


                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }

                    break;


                case "AMPLA":

                    var trdAMPLA = new Thread(() => varios(entrada)
                               );

                    trdAMPLA.IsBackground = true;
                    trdAMPLA.SetApartmentState(ApartmentState.STA);
                    trdAMPLA.Start();
                    while (trdAMPLA.IsAlive)
                    {

                    }


                    int arquivoAMPLA = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {


                        Console.WriteLine(Path.GetFileName(pdf));

                        int multiTiffID = 0;
                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int image = 1; int arq = 0;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {


                                string dirsaida = saida + @"\saida_.tif";

                                oGdPicturePDF.SelectPage(image);
                                image++;
                                int rasterizedPageID = oGdPicturePDF.RenderPageToGdPictureImageEx(300, true, pixelFormat);

                                multiTiffID = rasterizedPageID;
                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID, (byte)196);
                                status = oGdPictureImaging.TiffSaveAsMultiPageFile(multiTiffID, dirsaida, tiffCompression);
                                //release old image identifier
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageID);

                                oGdPictureImaging.TiffCloseMultiPageFile(multiTiffID);

                                oGdPictureImaging.ReleaseGdPictureImage(multiTiffID);




                                Image myImagesflip = Image.FromFile(dirsaida);
                                if (myImagesflip.Width < myImagesflip.Height)
                                {
                                    myImagesflip.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                    myImagesflip.Save(dirsaida);

                                }
                                myImagesflip.Dispose();
                                Image myImage = Image.FromFile(dirsaida);

                                Bitmap croppedBitmap = new Bitmap(myImage);
                                croppedBitmap.SetResolution(dpi, dpi);
                                croppedBitmap = croppedBitmap.Clone(
                                            new Rectangle(0, 0, ((myImage.Width) / 2), myImage.Height),
                                            pixelFormat);


                                int imgteste = oGdPictureImaging.CreateGdPictureImageFromBitmap(croppedBitmap);







                                oGdPictureImaging.ConvertTo1Bpp(imgteste);

                                oGdPictureImaging.SaveAsTIFF(imgteste, saida + @"\" + "_" + lote.ToString() + "_" + arq.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);

                                croppedBitmap.Dispose();




                                Bitmap croppedBitmaps = new Bitmap(myImage);
                                croppedBitmaps.SetResolution(dpi, dpi);


                                var d = (myImage.Width / 2) - 20;
                                croppedBitmaps = croppedBitmaps.Clone(
                                         new Rectangle(d, 0, (myImage.Width) / 2, myImage.Height),
                                            pixelFormat);
                                int imgteste2 = oGdPictureImaging.CreateGdPictureImageFromBitmap(croppedBitmaps);
                                myImage.Dispose();
                                oGdPictureImaging.ConvertTo1Bpp(imgteste2);

                                croppedBitmaps.Dispose();

                                var s = MergeTwoImagesXY(Image.FromFile(saida + @"\" + "_" + lote.ToString() + "_" + arq.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif"), 0, 0, Image.FromFile(saida + @"\" + "ROTULO.png"));




                                int cm = oGdPictureImaging.CreateGdPictureImageFromBitmap(s);

                                oGdPictureImaging.ConvertTo1Bpp(cm);

                                oGdPictureImaging.SaveAsTIFF(cm, saida + @"\" + "_" + lote.ToString() + "_" + arq.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);


                                s.Dispose();



                                arq++;
                                oGdPictureImaging.SaveAsTIFF(imgteste2, saida + @"\" + "_" + lote.ToString() + "_" + arq.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);

                                var sd = MergeTwoImagesXY(Image.FromFile(saida + @"\" + "_" + lote.ToString() + "_" + arq.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif"), 0, 0, Image.FromFile(saida + @"\" + "ROTULO.png"));
                                oGdPictureImaging.TiffCloseMultiPageFile(cm);
                                oGdPictureImaging.ReleaseGdPictureImage(cm);
                                oGdPictureImaging.ReleaseGdPictureImage(imgteste);
                                oGdPictureImaging.TiffCloseMultiPageFile(imgteste);

                                int cmi = oGdPictureImaging.CreateGdPictureImageFromBitmap(sd);

                                oGdPictureImaging.ConvertTo1Bpp(cmi);

                                oGdPictureImaging.SaveAsTIFF(cmi, saida + @"\" + "_" + lote.ToString() + "_" + arq.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", tiffCompression);
                                oGdPictureImaging.ReleaseGdPictureImage(cmi);
                                oGdPictureImaging.TiffCloseMultiPageFile(cmi);
                                sd.Dispose();




                                arq++;

                                oGdPictureImaging.ReleaseGdPictureImage(imgteste2);
                                File.Delete(dirsaida);


                            }



                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }
                    break;
                case "COELBA":
                    var trdCOELBA = new Thread(() => varios(entrada)
                     );

                    trdCOELBA.IsBackground = true;
                    trdCOELBA.SetApartmentState(ApartmentState.STA);
                    trdCOELBA.Start();
                    while (trdCOELBA.IsAlive)
                    {

                    }


                    int arquivoCOELBA = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        int multiTiffID = 0;
                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int image = 1;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {



                                var dirsaida = saida + @"\" + "_" + lote.ToString() + "_" + arquivoCOELBA.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";
                                arquivoCOELBA++;
                                oGdPicturePDF.SelectPage(image);
                                image++;
                                //render selected page to GdPictureImage identifier
                                int rasterizedPageID = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                multiTiffID = rasterizedPageID;

                                oGdPictureImaging.ConvertBitonalToGrayScale(rasterizedPageID, GrayScale);
                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID, 175);

                                //make the image a new multipage tiff image
                                status = oGdPictureImaging.TiffSaveAsMultiPageFile(multiTiffID, dirsaida, tiffCompression);
                                //release old image identifier
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageID);

                                oGdPictureImaging.TiffCloseMultiPageFile(multiTiffID);

                                oGdPictureImaging.ReleaseGdPictureImage(multiTiffID);

                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID);

                            }


                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }


                    break;
                case "COELCE":
                    var trdCOELCE = new Thread(() => varios(entrada)
                     );

                    trdCOELCE.IsBackground = true;
                    trdCOELCE.SetApartmentState(ApartmentState.STA);
                    trdCOELCE.Start();
                    while (trdCOELCE.IsAlive)
                    {

                    }


                    int arquivoCOELCE = 0;
                    int arquivoa = 0;
                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        int multiTiffID = 0;
                        //Load pdf file
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                        //if pdf loading was successful
                        if (status == GdPictureStatus.OK)
                        {
                            int image = 1;
                            //loop through pages
                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {


                                string dirsaida = saida + @"\" + "_" + lote.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";
                                arquivoCOELCE++;

                                oGdPicturePDF.SelectPage(image);
                                image++;
                                //render selected page to GdPictureImage identifier
                                int rasterizedPageID = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                //if it is the first page



                                //release old image identifier
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageID);



                                int rasterizedPageIDCrop = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                oGdPictureImaging.SaveAsTIFF(rasterizedPageIDCrop, saida + @"\" + "_" + "saidaCrop" + "_.tif", tiffCompression);
                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageIDCrop, 225);
                                oGdPictureImaging.SaveAsTIFF(rasterizedPageIDCrop, dirsaida, tiffCompression);
                                //release old image identifier
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageIDCrop);
                                oGdPictureImaging.TiffCloseMultiPageFile(rasterizedPageIDCrop);
                                Image myImagess = Image.FromFile(saida + @"\" + "_" + "saidaCrop" + "_.tif");





                                Bitmap croppedBitmapss = new Bitmap(myImagess);
                                croppedBitmapss.SetResolution(dpi, dpi);
                                croppedBitmapss = croppedBitmapss.Clone(
                                         new Rectangle(175, 5, 345, 705),
                                            pixelFormat);


                                int imgteste = oGdPictureImaging.CreateGdPictureImageFromBitmap(croppedBitmapss);

                                oGdPictureImaging.ConvertTo1Bpp(imgteste, (byte)163);

                                oGdPictureImaging.SaveAsTIFF(imgteste, saida + @"\CROP1.tif", tiffCompression);
                                oGdPictureImaging.ReleaseGdPictureImage(imgteste);

                                croppedBitmapss.Dispose();

                                Bitmap croppedBitmapCROP2 = new Bitmap(myImagess);
                                croppedBitmapCROP2.SetResolution(dpi, dpi);
                                croppedBitmapCROP2 = croppedBitmapCROP2.Clone(
                                         new Rectangle(1925, 7, 355, 701),
                                            pixelFormat);


                                int imgtestecroppedBitmapCROP2 = oGdPictureImaging.CreateGdPictureImageFromBitmap(croppedBitmapCROP2);

                                oGdPictureImaging.ConvertTo1Bpp(imgtestecroppedBitmapCROP2, (byte)163);

                                oGdPictureImaging.SaveAsTIFF(imgtestecroppedBitmapCROP2, saida + @"\CROP2.tif", tiffCompression);
                                oGdPictureImaging.ReleaseGdPictureImage(imgtestecroppedBitmapCROP2);
                                oGdPictureImaging.TiffCloseMultiPageFile(imgtestecroppedBitmapCROP2);
                                croppedBitmapCROP2.Dispose();



                                myImagess.Dispose();





                                multiTiffID = rasterizedPageID;
                                //make the image a new multipage tiff image




                                status = oGdPictureImaging.TiffSaveAsMultiPageFile(multiTiffID, dirsaida, tiffCompression);
                                //release old image identifier
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageID);



                                oGdPictureImaging.TiffCloseMultiPageFile(multiTiffID);

                                oGdPictureImaging.ReleaseGdPictureImage(multiTiffID);

                                ImageCodecInfo myImageCodecInfo;
                                myImageCodecInfo = GetEncoderInfo("image/tiff");
                                System.Drawing.Imaging.Encoder myEncoder;
                                myEncoder = System.Drawing.Imaging.Encoder.Compression;
                                EncoderParameters myEncoderParameters;
                                myEncoderParameters = new EncoderParameters(1);
                                EncoderParameter myEncoderParameter;
                                myEncoderParameter = new EncoderParameter(myEncoder, (long)EncoderValue.CompressionCCITT4);
                                myEncoderParameters.Param[0] = myEncoderParameter;




                                Image myImage = Image.FromFile(dirsaida);
                                Bitmap croppedBitmap = new Bitmap(myImage);
                                croppedBitmap = croppedBitmap.Clone(
                                            new Rectangle(0, 0, 1765, 2478),
                                            pixelFormat);
                                croppedBitmap.SetResolution(dpi, dpi);
                                var teste = @"\" + "_" + lote.ToString() + "_" + arquivoa.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";


                                string caminho1 = saida + @"\" + "_" + lote.ToString() + "_" + arquivoa.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";

                                croppedBitmap.Save(saida + @"\" + "_" + lote.ToString() + "_" + arquivoa.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", myImageCodecInfo, myEncoderParameters);
                                croppedBitmap.Dispose();
                                myImage.Dispose();








                                arquivoa++;
                                Image myImages = Image.FromFile(dirsaida);
                                Bitmap croppedBitmaps = new Bitmap(myImages);
                                croppedBitmaps = croppedBitmaps.Clone(
                                            new Rectangle(1777, 1, 1731, 2474),
                                            pixelFormat);
                                croppedBitmaps.SetResolution(dpi, dpi);
                                string caminho2 = saida + @"\" + "_" + lote.ToString() + "_" + arquivoa.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif";
                                croppedBitmaps.Save(saida + @"\" + "_" + lote.ToString() + "_" + arquivoa.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tif", myImageCodecInfo, myEncoderParameters);
                                arquivoa++;
                                croppedBitmaps.Dispose();
                                myImages.Dispose();
                                File.Delete(dirsaida);


                                var d = MergeTwoImagesXY(Image.FromFile(caminho1), 175, 0, Image.FromFile(saida + @"\" + "CROP1.tif"));
                                // d.Save(caminho1);


                                int cm = oGdPictureImaging.CreateGdPictureImageFromBitmap(d);

                                oGdPictureImaging.ConvertTo1Bpp(cm);

                                oGdPictureImaging.SaveAsTIFF(cm, caminho1, tiffCompression);
                                oGdPictureImaging.ReleaseGdPictureImage(cm);

                                d.Dispose();

                                File.Delete(saida + @"\" + "CROP1.tif");


                                var s = MergeTwoImagesXY(Image.FromFile(caminho2), 145, 0, Image.FromFile(saida + @"\" + "CROP2.tif"));
                                // s.Save(caminho2);



                                int gi = oGdPictureImaging.CreateGdPictureImageFromBitmap(s);

                                oGdPictureImaging.ConvertTo1Bpp(gi);

                                oGdPictureImaging.SaveAsTIFF(gi, caminho2, tiffCompression);
                                oGdPictureImaging.ReleaseGdPictureImage(gi);
                                s.Dispose();


                                File.Delete(saida + @"\" + "CROP2.tif");
                                File.Delete(saida + @"\" + "_saidaCrop_.tif");

                            }


                            oGdPicturePDF.CloseDocument();
                        }
                        moveFile(pdf);
                    }
                    break;

                case "ELEKTRO":
                    var trdELEKTRO = new Thread(() => varios(entrada));
                    trdELEKTRO.IsBackground = true;
                    trdELEKTRO.SetApartmentState(ApartmentState.STA);
                    trdELEKTRO.Start();
                    while (trdELEKTRO.IsAlive)
                    {

                    }

                    int arquivoELEKTRO = 0;

                    foreach (var pdf in Directory.GetFiles(entrada))
                    {
                        Console.WriteLine(Path.GetFileName(pdf));

                        int multiTiffID = 0;
                        int rasterizedPageID = 0;
                        int pag1 = 0;
                        GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdf, false);
                       
                        var dirsaidas = saida + @"\" + "_" + lote.ToString() + "_" + arquivoELEKTRO.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tiff";

                        if (status == GdPictureStatus.OK)
                        {

                            int image = 1;

                            for (int i = 0; i < oGdPicturePDF.GetPageCount(); i++)
                            {

                                var dirsaida = saida + @"\" + "_" + lote.ToString() + "_" + arquivoELEKTRO.ToString() + "_" + nomeUsuario + "_" + debito_nome + "_.tiff";
                                arquivoELEKTRO++;
                                oGdPicturePDF.SelectPage(image);
                                image++;

                                rasterizedPageID = oGdPicturePDF.RenderPageToGdPictureImageEx(dpi, true, pixelFormat);
                                multiTiffID = rasterizedPageID;

                                oGdPictureImaging.ConvertTo8BppGrayScaleAdv(rasterizedPageID);
                                oGdPictureImaging.ConvertTo1Bpp(rasterizedPageID, 158);
                                

                                if (image == 2)
                                {
                                    pag1 = multiTiffID;
                                    status = oGdPictureImaging.TiffSaveAsMultiPageFile(multiTiffID, dirsaida, tiffCompression);

                                }
                                if (status == GdPictureStatus.OK)
                                {


                                }
                                
                                if (image > 2)
                                {
                                    oGdPictureImaging.TiffAddToMultiPageFile(pag1, rasterizedPageID);
                                    oGdPictureImaging.ReleaseGdPictureImage(multiTiffID);
                                }
                                oGdPictureImaging.ReleaseGdPictureImage(rasterizedPageID);
                                oGdPictureImaging.ReleaseGdPictureImage(multiTiffID);

                            }
                            oGdPictureImaging.TiffCloseMultiPageFile(multiTiffID);

                            oGdPicturePDF.CloseDocument();
                        }

                        moveFile(pdf);
                    }
                    break;

            }
            if (Directory.GetFiles(saida).Count() > 0)
            {
                sendPutty(saida);

            }

        }

        public static Bitmap MergeTwoImagesXY(Image firstImage, int point1, int point2, Image secondImage)
        {
            if (firstImage == null)
            {
                throw new ArgumentNullException("firstImage");
            }

            if (secondImage == null)
            {
                throw new ArgumentNullException("secondImage");
            }


            ImageCodecInfo myImageCodecInfo;
            myImageCodecInfo = GetEncoderInfo("image/tiff");
            System.Drawing.Imaging.Encoder myEncoder;
            myEncoder = System.Drawing.Imaging.Encoder.Compression;
            EncoderParameters myEncoderParameters;
            myEncoderParameters = new EncoderParameters(1);
            EncoderParameter myEncoderParameter;
            myEncoderParameter = new EncoderParameter(myEncoder, (long)EncoderValue.CompressionCCITT4);
            myEncoderParameters.Param[0] = myEncoderParameter;

            Bitmap outputImage = new Bitmap(firstImage.Width, firstImage.Height);
            outputImage.SetResolution(300, 300);
            using (Graphics graphics = Graphics.FromImage(outputImage))
            {
                graphics.DrawImage(firstImage, new Point(0, 0));
                graphics.DrawImage(secondImage, new Point(point1, point2));
            }
            firstImage.Dispose();
            secondImage.Dispose();
            return outputImage;
        }

        public static Bitmap MergeTwoImages(Image firstImage, Image secondImage)
        {
            if (firstImage == null)
            {
                throw new ArgumentNullException("firstImage");
            }

            if (secondImage == null)
            {
                throw new ArgumentNullException("secondImage");
            }

            int outputImageWidth = firstImage.Width > secondImage.Width ? firstImage.Width : secondImage.Width;

            int outputImageHeight = firstImage.Height + secondImage.Height + 1;

            ImageCodecInfo myImageCodecInfo;
            myImageCodecInfo = GetEncoderInfo("image/tiff");
            System.Drawing.Imaging.Encoder myEncoder;
            myEncoder = System.Drawing.Imaging.Encoder.Compression;
            EncoderParameters myEncoderParameters;
            myEncoderParameters = new EncoderParameters(1);
            EncoderParameter myEncoderParameter;
            myEncoderParameter = new EncoderParameter(myEncoder, (long)EncoderValue.CompressionCCITT4);
            myEncoderParameters.Param[0] = myEncoderParameter;

            Bitmap outputImage = new Bitmap(outputImageWidth, outputImageHeight);
            outputImage.SetResolution(300, 300);
            using (Graphics graphics = Graphics.FromImage(outputImage))
            {
                graphics.DrawImage(firstImage, new Rectangle(new Point(), firstImage.Size),
                    new Rectangle(new Point(), firstImage.Size), GraphicsUnit.Pixel);
                graphics.DrawImage(secondImage, new Rectangle(new Point(0, firstImage.Height + 1), secondImage.Size),
                    new Rectangle(new Point(), secondImage.Size), GraphicsUnit.Pixel);
            }
            firstImage.Dispose();
            secondImage.Dispose();
            return outputImage;
        }


        public static void sendPutty(String pastaSaida)
        {

            String usuario = ConfigurationManager.AppSettings["usuario"].ToString();

            String Senha = ConfigurationManager.AppSettings["Senha"].ToString();

            String Servidor = ConfigurationManager.AppSettings["Servidor"].ToString();

            var caminhoPutty = ConfigurationManager.AppSettings["Putty"].ToString();
            var ComandoPutty = ConfigurationManager.AppSettings["Comando"].ToString();

            var arquivos = Directory.GetFiles(pastaSaida).ToList();


            Console.WriteLine("ENtrando no send");
            var processStartInfo = new ProcessStartInfo();
            processStartInfo.WorkingDirectory = caminhoPutty;
            processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            processStartInfo.FileName = ComandoPutty;
            processStartInfo.Arguments = " -q -pw " + Senha + " " + pastaSaida + @"\*" + " " + usuario + Servidor + ":./";
            // set additional properties     

            //   Console.WriteLine(caminhoPutty+ " -q -pw " + Senha + " " + item + " " + usuario + Servidor + ":./" + Path.GetFileName(item));
            // Console.ReadLine();
            try
            {
                Process proc = Process.Start(processStartInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            //var proc = new Process()
            //{
            //    StartInfo = new ProcessStartInfo
            //    {
            //        FileName = caminhoPutty,
            //        Arguments = "-q - pw " + Senha + " " + item + " " + usuario + Servidor + ":./ " + Path.GetFileName(item),
            //        WindowStyle = ProcessWindowStyle.Hidden
            //    }
            //};
            //proc.Start();




        }

        private static ImageCodecInfo GetEncoderInfo(String mimeType)
        {
            int j;
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (j = 0; j < encoders.Length; ++j)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }
            return null;
        }
    }

}

