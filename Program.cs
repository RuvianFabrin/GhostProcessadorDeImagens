using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ImageMagick;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace ProcessadorDeImagens
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("=== Menu ===");
                Console.WriteLine("1 - Converter imagens para .jpg");
                Console.WriteLine("2 - Redimensionar imagens");
                Console.WriteLine("3 - Colar imagens");
                Console.WriteLine("4 - Converter imagens em PDF");
                Console.WriteLine("5 - Fazer todos");
                Console.WriteLine("6 - Sair");
                Console.Write("Selecione uma opção: ");
                string opcao = Console.ReadLine();

                switch (opcao)
                {
                    case "1":
                        ConverterImagensParaJpg();
                        break;
                    case "2":
                        RedimensionarImagens();
                        break;
                    case "3":
                        ColarImagens();
                        break;
                    case "4":
                        ConverterImagensParaPdf();
                        break;
                    case "5":
                        FazerTodos();
                        break;
                    case "6":
                        return;
                    default:
                        Console.WriteLine("Opção inválida. Pressione qualquer tecla para tentar novamente...");
                        Console.ReadKey();
                        break;
                }
            }
        }

        // Implementação da ordenação natural
        public class NaturalStringComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                return StrCmpLogicalW(x, y);
            }

            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
            private static extern int StrCmpLogicalW(string psz1, string psz2);
        }

        static void ConverterImagensParaJpg()
        {
            Console.Write("Informe o caminho da pasta onde estão as imagens: ");
            string caminhoPasta = Console.ReadLine();

            if (!Directory.Exists(caminhoPasta))
            {
                Console.WriteLine("A pasta informada não existe. Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            string pastaDestino = Path.Combine(caminhoPasta, "imagens_convertidas");
            Directory.CreateDirectory(pastaDestino);

            string[] arquivosImagem = Directory.GetFiles(caminhoPasta, "*.*", SearchOption.TopDirectoryOnly);
            string[] extensoesPermitidas = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".ico", ".webp" };

            foreach (string arquivo in arquivosImagem)
            {
                if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                {
                    string nomeArquivo = Path.GetFileNameWithoutExtension(arquivo);
                    string caminhoDestino = Path.Combine(pastaDestino, nomeArquivo + ".jpg");

                    try
                    {
                        using (MagickImage image = new MagickImage(arquivo))
                        {
                            image.Format = MagickFormat.Jpeg;
                            image.Write(caminhoDestino);
                            Console.WriteLine($"Imagem '{nomeArquivo}' convertida para JPG.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao processar a imagem '{nomeArquivo}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine("Processo concluído. Pressione qualquer tecla para voltar ao menu...");
            Console.ReadKey();
        }

        static void RedimensionarImagens()
        {
            Console.Write("Informe o caminho da pasta onde estão as imagens: ");
            string caminhoPasta = Console.ReadLine();

            if (!Directory.Exists(caminhoPasta))
            {
                Console.WriteLine("A pasta informada não existe. Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            string pastaDestino = Path.Combine(caminhoPasta, "imagens_redimensionadas");
            Directory.CreateDirectory(pastaDestino);

            string[] arquivosImagem = Directory.GetFiles(caminhoPasta, "*.*", SearchOption.TopDirectoryOnly);
            string[] extensoesPermitidas = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".ico", ".webp" };

            foreach (string arquivo in arquivosImagem)
            {
                if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                {
                    string nomeArquivo = Path.GetFileName(arquivo);
                    string caminhoDestino = Path.Combine(pastaDestino, nomeArquivo);

                    try
                    {
                        using (MagickImage image = new MagickImage(arquivo))
                        {
                            uint larguraAtual = image.Width;
                            uint novaLargura = 700;
                            uint novaAltura = (uint)(image.Height * ((double)novaLargura / larguraAtual));

                            if (larguraAtual == novaLargura)
                            {
                                image.Write(caminhoDestino);
                                Console.WriteLine($"Imagem '{nomeArquivo}' já está em 700px, copiada sem alteração.");
                            }
                            else
                            {
                                // Verifica se os valores estão dentro dos limites de int
                                if (novaLargura > int.MaxValue || novaAltura > int.MaxValue)
                                {
                                    Console.WriteLine($"Dimensões muito grandes para a imagem '{nomeArquivo}'.");
                                    continue;
                                }

                                image.Resize(novaLargura, novaAltura);
                                image.Write(caminhoDestino);
                                Console.WriteLine($"Imagem '{nomeArquivo}' redimensionada.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao processar a imagem '{nomeArquivo}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine("Processo concluído. Pressione qualquer tecla para voltar ao menu...");
            Console.ReadKey();
        }

        static void ColarImagens()
        {
            Console.Write("Informe o caminho da pasta onde estão as imagens: ");
            string caminhoPasta = Console.ReadLine();

            if (!Directory.Exists(caminhoPasta))
            {
                Console.WriteLine("A pasta informada não existe. Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            string pastaGrande = Path.Combine(caminhoPasta, "imagem_grande");
            Directory.CreateDirectory(pastaGrande);

            string pastaDestino = Path.Combine(caminhoPasta, "imagens_cortadas");
            Directory.CreateDirectory(pastaDestino);

            string[] arquivosImagem = Directory.GetFiles(caminhoPasta, "*.*", SearchOption.TopDirectoryOnly);
            string[] extensoesPermitidas = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".ico", ".webp" };

            List<string> imagensParaProcessar = new List<string>();

            // Filtra apenas os arquivos de imagem suportados e ordena naturalmente
            foreach (string arquivo in arquivosImagem)
            {
                if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                {
                    imagensParaProcessar.Add(arquivo);
                }
            }

            // Ordenação natural dos arquivos
            imagensParaProcessar.Sort(new NaturalStringComparer());

            if (imagensParaProcessar.Count == 0)
            {
                Console.WriteLine("Nenhuma imagem encontrada na pasta especificada.");
                Console.WriteLine("Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            List<MagickImage> imagensParaConcatenar = new List<MagickImage>();
            uint larguraMaxima = 0;
            uint alturaTotal = 0;

            foreach (string caminhoImagem in imagensParaProcessar)
            {
                try
                {
                    MagickImage img = new MagickImage(caminhoImagem);
                    imagensParaConcatenar.Add(img);
                    larguraMaxima = Math.Max(larguraMaxima, img.Width);
                    alturaTotal += img.Height;
                    Console.WriteLine($"Imagem '{Path.GetFileName(caminhoImagem)}' adicionada para concatenação.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao carregar a imagem '{Path.GetFileName(caminhoImagem)}': {ex.Message}");
                }
            }

            if (imagensParaConcatenar.Count == 0)
            {
                Console.WriteLine("Nenhuma imagem para concatenar.");
                Console.WriteLine("Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            // Verifica se larguraMaxima e alturaTotal não excedem int.MaxValue
            if (larguraMaxima > int.MaxValue || alturaTotal > int.MaxValue)
            {
                Console.WriteLine("Dimensões da imagem grande excedem o limite suportado.");
                Console.WriteLine("Processo não pode ser concluído.");
                Console.ReadKey();
                return;
            }

            // Define as configurações para a nova imagem
            MagickReadSettings settings = new MagickReadSettings()
            {
                BackgroundColor = MagickColors.Black,
                Width = larguraMaxima,
                Height = alturaTotal
            };

            // Cria a imagem com base nas configurações
            using (MagickImage imagemFinal = new MagickImage("xc:black", settings))
            {
                long posicaoY = 0;

                foreach (MagickImage img in imagensParaConcatenar)
                {
                    imagemFinal.Composite(img, 0, (int)posicaoY, CompositeOperator.Over);
                    posicaoY += img.Height;
                    img.Dispose(); // Libera a imagem após o uso
                }

                string caminhoImagemGrande = Path.Combine(pastaGrande, "imagem_grande.png");
                imagemFinal.Write(caminhoImagemGrande);
                Console.WriteLine($"Imagem grande salva em '{caminhoImagemGrande}'.");

                // Cortar a imagem grande em pedaços de 10.000px de altura
                int numPecas = (int)Math.Ceiling((double)alturaTotal / 10000);

                for (int i = 0; i < numPecas; i++)
                {
                    uint alturaCorte = (i == numPecas - 1) ? (uint)(alturaTotal - (i * 10000)) : 10000;
                    MagickGeometry areaCorte = new MagickGeometry(0, i * 10000, larguraMaxima, alturaCorte);

                    using (IMagickImage<byte> imagemCortada = imagemFinal.Clone(areaCorte))
                    {
                        string nomeArquivo = $"imagem_cortada_{i + 1}.png";
                        string caminhoDestino = Path.Combine(pastaDestino, nomeArquivo);
                        imagemCortada.Write(caminhoDestino);
                        Console.WriteLine($"Imagem cortada '{nomeArquivo}' salva com sucesso.");
                    }
                }
            }

            Console.WriteLine("Processo concluído. Pressione qualquer tecla para voltar ao menu...");
            Console.ReadKey();
        }

        static void ConverterImagensParaPdf()
        {
            Console.Write("Informe o caminho da pasta onde estão as imagens: ");
            string caminhoPasta = Console.ReadLine();

            if (!Directory.Exists(caminhoPasta))
            {
                Console.WriteLine("A pasta informada não existe. Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            string pastaDestino = Path.Combine(caminhoPasta, "pdf");
            Directory.CreateDirectory(pastaDestino);

            string[] arquivosImagem = Directory.GetFiles(caminhoPasta, "*.*", SearchOption.TopDirectoryOnly);
            string[] extensoesPermitidas = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".ico", ".webp" };

            List<string> imagensParaProcessar = new List<string>();

            // Filtra apenas os arquivos de imagem suportados e ordena naturalmente
            foreach (string arquivo in arquivosImagem)
            {
                if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                {
                    imagensParaProcessar.Add(arquivo);
                }
            }

            // Ordenação natural dos arquivos
            imagensParaProcessar.Sort(new NaturalStringComparer());

            if (imagensParaProcessar.Count == 0)
            {
                Console.WriteLine("Nenhuma imagem encontrada na pasta especificada.");
                Console.WriteLine("Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            PdfDocument documento = new PdfDocument();

            foreach (string caminhoImagem in imagensParaProcessar)
            {
                try
                {
                    using (MagickImage image = new MagickImage(caminhoImagem))
                    {
                        // Define a densidade da imagem para 72 DPI
                        image.Density = new Density(72, 72);

                        // Converte para RGB se necessário
                        if (image.ColorSpace != ColorSpace.RGB)
                        {
                            image.ColorSpace = ColorSpace.RGB;
                        }

                        double width = image.Width;
                        double height = image.Height;

                        PdfPage pagina = documento.AddPage();
                        pagina.Width = width;
                        pagina.Height = height;

                        using (XGraphics gfx = XGraphics.FromPdfPage(pagina))
                        {
                            // Salva a imagem em um arquivo temporário
                            string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                            image.Write(tempImagePath);

                            using (XImage xImage = XImage.FromFile(tempImagePath))
                            {
                                gfx.DrawImage(xImage, 0, 0, width, height);
                            }

                            // Apaga o arquivo temporário
                            File.Delete(tempImagePath);
                        }

                        Console.WriteLine($"Imagem '{Path.GetFileName(caminhoImagem)}' adicionada ao PDF.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao processar a imagem '{Path.GetFileName(caminhoImagem)}': {ex.Message}");
                }
            }

            string caminhoDestino = Path.Combine(pastaDestino, "documento.pdf");
            documento.Save(caminhoDestino);
            documento.Close();

            Console.WriteLine($"PDF criado e salvo em '{caminhoDestino}'.");
            Console.WriteLine("Processo concluído. Pressione qualquer tecla para voltar ao menu...");
            Console.ReadKey();
        }

        static void FazerTodos()
        {
            Console.Write("Informe o caminho da pasta principal (que contém as subpastas): ");
            string caminhoPastaPrincipal = Console.ReadLine();

            if (!Directory.Exists(caminhoPastaPrincipal))
            {
                Console.WriteLine("A pasta informada não existe. Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            string[] subpastas = Directory.GetDirectories(caminhoPastaPrincipal);

            if (subpastas.Length == 0)
            {
                Console.WriteLine("Nenhuma subpasta encontrada na pasta principal.");
                Console.WriteLine("Pressione qualquer tecla para voltar ao menu...");
                Console.ReadKey();
                return;
            }

            foreach (string subpasta in subpastas)
            {
                Console.WriteLine($"\nProcessando a subpasta '{Path.GetFileName(subpasta)}'...");

                // Passo 1: Converter imagens para .jpg
                string pastaConvertidas = Path.Combine(subpasta, "imagens_convertidas");
                Directory.CreateDirectory(pastaConvertidas);

                string[] arquivosImagem = Directory.GetFiles(subpasta, "*.*", SearchOption.TopDirectoryOnly);
                string[] extensoesPermitidas = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".ico", ".webp" };

                foreach (string arquivo in arquivosImagem)
                {
                    if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                    {
                        string nomeArquivo = Path.GetFileNameWithoutExtension(arquivo);
                        string caminhoDestino = Path.Combine(pastaConvertidas, nomeArquivo + ".jpg");

                        try
                        {
                            using (MagickImage image = new MagickImage(arquivo))
                            {
                                image.Format = MagickFormat.Jpeg;
                                image.Write(caminhoDestino);
                                Console.WriteLine($"Imagem '{nomeArquivo}' convertida para JPG.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erro ao processar a imagem '{nomeArquivo}': {ex.Message}");
                        }
                    }
                }

                // Passo 2: Redimensionar imagens
                string pastaRedimensionadas = Path.Combine(subpasta, "imagens_redimensionadas");
                Directory.CreateDirectory(pastaRedimensionadas);

                string[] imagensConvertidas = Directory.GetFiles(pastaConvertidas, "*.*", SearchOption.TopDirectoryOnly);

                foreach (string arquivo in imagensConvertidas)
                {
                    if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                    {
                        string nomeArquivo = Path.GetFileName(arquivo);
                        string caminhoDestino = Path.Combine(pastaRedimensionadas, nomeArquivo);

                        try
                        {
                            using (MagickImage image = new MagickImage(arquivo))
                            {
                                uint larguraAtual = image.Width;
                                uint novaLargura = 700;
                                uint novaAltura = (uint)(image.Height * ((double)novaLargura / larguraAtual));

                                if (larguraAtual == novaLargura)
                                {
                                    image.Write(caminhoDestino);
                                    Console.WriteLine($"Imagem '{nomeArquivo}' já está em 700px, copiada sem alteração.");
                                }
                                else
                                {
                                    if (novaLargura > int.MaxValue || novaAltura > int.MaxValue)
                                    {
                                        Console.WriteLine($"Dimensões muito grandes para a imagem '{nomeArquivo}'.");
                                        continue;
                                    }

                                    image.Resize(novaLargura, novaAltura);
                                    image.Write(caminhoDestino);
                                    Console.WriteLine($"Imagem '{nomeArquivo}' redimensionada.");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erro ao processar a imagem '{nomeArquivo}': {ex.Message}");
                        }
                    }
                }

                // Passo 3: Colar imagens
                string pastaGrande = Path.Combine(subpasta, "imagem_grande");
                Directory.CreateDirectory(pastaGrande);

                string pastaCortadas = Path.Combine(subpasta, "imagens_cortadas");
                Directory.CreateDirectory(pastaCortadas);

                string[] imagensRedimensionadas = Directory.GetFiles(pastaRedimensionadas, "*.*", SearchOption.TopDirectoryOnly);

                List<string> imagensParaProcessar = new List<string>();

                foreach (string arquivo in imagensRedimensionadas)
                {
                    if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                    {
                        imagensParaProcessar.Add(arquivo);
                    }
                }

                // Ordenação natural dos arquivos
                imagensParaProcessar.Sort(new NaturalStringComparer());

                if (imagensParaProcessar.Count == 0)
                {
                    Console.WriteLine("Nenhuma imagem para concatenar.");
                }
                else
                {
                    List<MagickImage> imagensParaConcatenar = new List<MagickImage>();
                    uint larguraMaxima = 0;
                    uint alturaTotal = 0;

                    foreach (string caminhoImagem in imagensParaProcessar)
                    {
                        try
                        {
                            MagickImage img = new MagickImage(caminhoImagem);
                            imagensParaConcatenar.Add(img);
                            larguraMaxima = Math.Max(larguraMaxima, img.Width);
                            alturaTotal += img.Height;
                            Console.WriteLine($"Imagem '{Path.GetFileName(caminhoImagem)}' adicionada para concatenação.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erro ao carregar a imagem '{Path.GetFileName(caminhoImagem)}': {ex.Message}");
                        }
                    }

                    // Verifica se larguraMaxima e alturaTotal não excedem int.MaxValue
                    if (larguraMaxima > int.MaxValue || alturaTotal > int.MaxValue)
                    {
                        Console.WriteLine("Dimensões da imagem grande excedem o limite suportado.");
                        Console.WriteLine("Processo não pode ser concluído.");
                        Console.ReadKey();
                        continue;
                    }

                    // Define as configurações para a nova imagem
                    MagickReadSettings settings = new MagickReadSettings()
                    {
                        BackgroundColor = MagickColors.Black,
                        Width = larguraMaxima,
                        Height = alturaTotal
                    };

                    // Cria a imagem com base nas configurações
                    using (MagickImage imagemFinal = new MagickImage("xc:black", settings))
                    {
                        long posicaoY = 0;

                        foreach (MagickImage img in imagensParaConcatenar)
                        {
                            imagemFinal.Composite(img, 0, (int)posicaoY, CompositeOperator.Over);
                            posicaoY += img.Height;
                            img.Dispose(); // Libera a imagem após o uso
                        }

                        string caminhoImagemGrande = Path.Combine(pastaGrande, "imagem_grande.png");
                        imagemFinal.Write(caminhoImagemGrande);
                        Console.WriteLine($"Imagem grande salva em '{caminhoImagemGrande}'.");

                        // Cortar a imagem grande em pedaços de 10.000px de altura
                        int numPecas = (int)Math.Ceiling((double)alturaTotal / 10000);

                        for (int i = 0; i < numPecas; i++)
                        {
                            uint alturaCorte = (i == numPecas - 1) ? (uint)(alturaTotal - (i * 10000)) : 10000;
                            MagickGeometry areaCorte = new MagickGeometry(0, i * 10000, larguraMaxima, alturaCorte);

                            using (IMagickImage<byte> imagemCortada = imagemFinal.Clone(areaCorte))
                            {
                                string nomeArquivo = $"imagem_cortada_{i + 1}.png";
                                string caminhoDestino = Path.Combine(pastaCortadas, nomeArquivo);
                                imagemCortada.Write(caminhoDestino);
                                Console.WriteLine($"Imagem cortada '{nomeArquivo}' salva com sucesso.");
                            }
                        }
                    }
                }

                // Passo 4: Converter imagens em PDF
                string pastaPdf = Path.Combine(pastaCortadas, "pdf");
                Directory.CreateDirectory(pastaPdf);

                string[] imagensParaPdf = Directory.GetFiles(pastaCortadas, "*.*", SearchOption.TopDirectoryOnly);

                List<string> imagensParaPdfProcessar = new List<string>();

                foreach (string arquivo in imagensParaPdf)
                {
                    if (Array.Exists(extensoesPermitidas, ext => ext == Path.GetExtension(arquivo).ToLower()))
                    {
                        imagensParaPdfProcessar.Add(arquivo);
                    }
                }

                // Ordenação natural dos arquivos
                imagensParaPdfProcessar.Sort(new NaturalStringComparer());

                if (imagensParaPdfProcessar.Count == 0)
                {
                    Console.WriteLine("Nenhuma imagem para converter em PDF.");
                }
                else
                {
                    PdfDocument documento = new PdfDocument();

                    foreach (string caminhoImagem in imagensParaPdfProcessar)
                    {
                        try
                        {
                            using (MagickImage image = new MagickImage(caminhoImagem))
                            {
                                // Define a densidade da imagem para 72 DPI
                                image.Density = new Density(72, 72);

                                // Converte para RGB se necessário
                                if (image.ColorSpace != ColorSpace.RGB)
                                {
                                    image.ColorSpace = ColorSpace.RGB;
                                }

                                double width = image.Width;
                                double height = image.Height;

                                PdfPage pagina = documento.AddPage();
                                pagina.Width = width;
                                pagina.Height = height;

                                using (XGraphics gfx = XGraphics.FromPdfPage(pagina))
                                {
                                    // Salva a imagem em um arquivo temporário
                                    string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                                    image.Write(tempImagePath);

                                    using (XImage xImage = XImage.FromFile(tempImagePath))
                                    {
                                        gfx.DrawImage(xImage, 0, 0, width, height);
                                    }

                                    // Apaga o arquivo temporário
                                    File.Delete(tempImagePath);
                                }

                                Console.WriteLine($"Imagem '{Path.GetFileName(caminhoImagem)}' adicionada ao PDF.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erro ao processar a imagem '{Path.GetFileName(caminhoImagem)}': {ex.Message}");
                        }
                    }

                    string caminhoPdfDestino = Path.Combine(pastaPdf, "documento.pdf");
                    documento.Save(caminhoPdfDestino);
                    documento.Close();

                    Console.WriteLine($"PDF criado e salvo em '{caminhoPdfDestino}'.");
                    Console.WriteLine($"Processamento da subpasta '{Path.GetFileName(pastaCortadas)}' concluído.");
                }
            }

            Console.WriteLine("\nProcesso concluído para todas as subpastas. Pressione qualquer tecla para voltar ao menu...");
            Console.ReadKey();
        }
    }
}
