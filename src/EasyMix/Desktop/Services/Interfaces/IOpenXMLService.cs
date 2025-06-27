using Desktop.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Desktop.Services.Interfaces
{
    public interface IOpenXMLService
    {
        Task<WordprocessingDocument> CreateDocumentAsync(string filePath);
        Task<WordprocessingDocument> OpenDocumentAsync(string filePath, bool isEditTable);
        Task SaveDocumentAsync(WordprocessingDocument document);
        Task CloseDocumentAsync(WordprocessingDocument document);
        Task<string> GetDocumentTextAsync(WordprocessingDocument document);
        Task<Body> GetDocumentBodyAsync(WordprocessingDocument document);
        Task FormatAllParagraphsAsync(WordprocessingDocument doc);

        Task GenerateShuffledExamsAsync(string sourceFile, string outputFolder, MixInfo mixInfo);
        Task<List<Question>> ParseDocxQuestionsAsync(string filePath);
    }
}
