using System;
using System.Windows.Forms;
using System.Net.Http;
using System.Text;
using System.Data.SqlClient;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace Proyecto_l
{
    public partial class Form1 : Form
    {
        private readonly string connectionString = "Server=DESKTOP-OQRPO5C\\SQLEXPRESS01;Database=InvestigacionDB;Trusted_Connection=True;";
        private readonly string apiKey = "";

        public Form1()
        {
            InitializeComponent();
            progressBar1.Visible = false;
        }

        private async void btnConsultar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPrompt.Text))
            {
                MessageBox.Show("Por favor, ingrese un prompt.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                btnConsultar.Enabled = false;
                txtResultado.Enabled = false;
                btnAprobar.Enabled = false;
                btnEditar.Enabled = false;

                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.MarqueeAnimationSpeed = 60;

                string resultado = await ConsultarAPI(txtPrompt.Text);
                txtResultado.Text = resultado;

                btnAprobar.Enabled = true;
                btnEditar.Enabled = true;
                txtResultado.ReadOnly = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al consultar la API: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar1.Visible = false;
                progressBar1.Style = ProgressBarStyle.Blocks;
                btnConsultar.Enabled = true;
                txtResultado.Enabled = true;
            }
        }

        private async Task<string> ConsultarAPI(string prompt)
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                client.DefaultRequestHeaders.Add("Accept", "application/json");

                var payload = new
                {
                    model = "meta/llama-4-maverick-17b-128e-instruct",
                    messages = new[]
                    {
                        new { role = "user", content =$"Quiero que hagas una investigacion corta sobre el siguiente titulo :  {prompt}, no me des ningun saludo ni despedida se directo y hazlo de manera academica" }
                    },
                    max_tokens = 2048,
                    temperature = 1.00,
                    top_p = 1.00,
                    stream = false
                };

                var content = new StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json");
                var response = await client.PostAsync("https://integrate.api.nvidia.com/v1/chat/completions", content);

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception($"Error en la API: {response.StatusCode} - {await response.Content.ReadAsStringAsync()}");
                }

                var responseData = await response.Content.ReadAsStringAsync();
                dynamic jsonResponse = JsonConvert.DeserializeObject(responseData);
                string result = jsonResponse.choices[0].message.content.ToString();
                if (string.IsNullOrWhiteSpace(result))
                {
                    throw new Exception("La API devolvió una respuesta vacía.");
                }
                return result;
            }
        }

        private void btnAprobar_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtResultado.Text))
            {
                MessageBox.Show("No hay resultado para guardar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                btnAprobar.Enabled = false;
                btnEditar.Enabled = false;
                btnConsultar.Enabled = false;

                progressBar1.Visible = true;
                progressBar1.Style = ProgressBarStyle.Continuous;
                progressBar1.Value = 0;

                progressBar1.Value = 20;
                GuardarEnBaseDatos(txtPrompt.Text, txtResultado.Text);

                progressBar1.Value = 40;
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"Investigaciones_{timestamp}");
                try
                {
                    Directory.CreateDirectory(folderPath);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error al crear la carpeta: {ex.Message}");
                }

                progressBar1.Value = 60;
                string wordPath = GenerarDocumentoWord(txtPrompt.Text, txtResultado.Text, folderPath);

                progressBar1.Value = 80;
                string pptPath = GenerarPresentacionPowerPoint(txtPrompt.Text, txtResultado.Text, folderPath);

                progressBar1.Value = 100;

                MessageBox.Show($"Consulta guardada y documentos generados en {folderPath}.\nWord: {Path.GetFileName(wordPath)}\nPowerPoint: {Path.GetFileName(pptPath)}",
                    "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                txtPrompt.Clear();
                txtResultado.Clear();
                btnAprobar.Enabled = false;
                btnEditar.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar1.Visible = false;
                progressBar1.Value = 0;
                btnConsultar.Enabled = true;
            }
        }

        private void btnEditar_Click(object sender, EventArgs e)
        {
            txtResultado.ReadOnly = false;
            txtResultado.Focus();
        }

        private void GuardarEnBaseDatos(string prompt, string resultado)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = "INSERT INTO Consultas (Prompt, Resultado) VALUES (@Prompt, @Resultado)";
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Prompt", prompt);
                        cmd.Parameters.AddWithValue("@Resultado", resultado);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al guardar en la base de datos: {ex.Message}");
            }
        }

        private string GenerarDocumentoWord(string prompt, string resultado, string folderPath)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;
            string filePath = string.Empty;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                doc = wordApp.Documents.Add();

                try
                {
                    string logoPath = @"C:\Users\Jose Carlos\OneDrive\Documentos\imagendelau\image.png";
                    if (File.Exists(logoPath))
                    {
                        var header = doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                        header.Range.Select();
                        var logo = wordApp.Selection.InlineShapes.AddPicture(logoPath);
                        logo.Height = 50;
                        logo.Width = 100;
                    }
                    else
                    {
                        throw new Exception($"No se encontró el logo en: {logoPath}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al agregar el logo en Word: {ex.Message}");
                }

                Word.Paragraph titleParagraph = doc.Content.Paragraphs.Add();
                titleParagraph.Range.Text = prompt;
                titleParagraph.Range.Font.Bold = 1;
                titleParagraph.Range.Font.Size = 16;
                titleParagraph.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                titleParagraph.Range.InsertParagraphAfter();

                Word.Paragraph contentParagraph = doc.Content.Paragraphs.Add();
                contentParagraph.Range.Text = resultado;
                contentParagraph.Range.Font.Size = 12;
                contentParagraph.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;

                filePath = Path.Combine(folderPath, $"Investigacion_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.docx");
                doc.SaveAs2(filePath);
                return filePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al generar el documento de Word: {ex.Message}");
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private string GenerarPresentacionPowerPoint(string prompt, string resultado, string folderPath)
        {
            PowerPoint.Application pptApp = null;
            PowerPoint.Presentation presentation = null;
            string filePath = string.Empty;

            try
            {
                pptApp = new PowerPoint.Application();
                presentation = pptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);

                PowerPoint.Slide slide1 = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);

                if (slide1.Shapes.Title != null)
                {
                    slide1.Shapes.Title.TextFrame.TextRange.Text = prompt;
                    slide1.Shapes.Title.TextFrame.TextRange.Font.Size = 45;
                    slide1.Shapes.Title.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                    slide1.Shapes.Title.TextFrame.HorizontalAnchor = Microsoft.Office.Core.MsoHorizontalAnchor.msoAnchorCenter;
                    

                }

                try
                {
                    string logoPath = @"C:\Users\Jose Carlos\OneDrive\Documentos\imagendelau\image.png";
                    if (File.Exists(logoPath))
                    {
                        float logoLeft = 20;
                        float logoTop = 20;
                        float logoWidth = 100;
                        float logoHeight = 50;

                        var logo = slide1.Shapes.AddPicture(
                            logoPath,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            logoLeft, logoTop, logoWidth, logoHeight);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error al agregar el logo en PowerPoint: {ex.Message}");
                }

                string[] parrafos = resultado.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);

                int maxSlides = Math.Min(parrafos.Length, 5);

                for (int i = 0; i < maxSlides; i++)
                {
                    PowerPoint.Slide slide = presentation.Slides.Add(i + 2, PowerPoint.PpSlideLayout.ppLayoutText);

                    if (slide.Shapes.Title != null)
                    {
                        slide.Shapes.Title.TextFrame.TextRange.Text = $"Sección {i + 1}";
                        slide.Shapes.Title.TextFrame.TextRange.Font.Size = 24;
                        slide.Shapes.Title.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                    }

                    PowerPoint.Shape contentShape = null;
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape != slide.Shapes.Title && shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            contentShape = shape;
                            break;
                        }
                    }

                    if (contentShape != null)
                    {
                        contentShape.TextFrame.TextRange.Text = parrafos[i].Trim();
                        contentShape.TextFrame.TextRange.Font.Size = 18;
                    }
                }

                filePath = Path.Combine(folderPath, $"Presentacion_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.pptx");
                presentation.SaveAs(filePath);
                return filePath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al generar la presentación de PowerPoint: {ex.Message}");
            }
            finally
            {
                if (presentation != null)
                {
                    presentation.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation);
                }
                if (pptApp != null)
                {
                    pptApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pptApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}