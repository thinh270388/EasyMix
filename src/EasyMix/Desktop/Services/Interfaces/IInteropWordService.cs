using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using Task = System.Threading.Tasks.Task;

namespace Desktop.Services.Interfaces
{
    public interface IInteropWordService
    {
        Task<Document> OpenDocumentAsync(string filePath, bool visible);
        Task SaveDocumentAsync(_Document document);
        Task CloseDocumentAsync(_Document document);
        ValueTask DisposeAsync();
        Task FormatDocumentAsync(_Document document);
        Task FindAndReplaceAsync(_Document document, string findText, string replaceWithText, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false);
        Task FindAndReplaceAsync(_Document document, Dictionary<string, string> replacements, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false);
        Task FindAndReplaceFirstAsync(Paragraph paragraph, string findText, string replaceWithText, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false);
        Task<string> ConvertDocxToXpsAsync(string docxPath);
        Task UpdateFieldsAsync(string filePath);
        Task SetAnswersToABCDAsync(_Document document);
        Task SetQuestionsToNumberAsync(_Document document);
        Task FormatQuestionAndAnswerAsync(_Document document);
        Task ProcessImagesInDocumentAsync(_Document document, bool isBorderImage);
        Task FindAndReplaceRedToUnderlinedAsync(_Document document);
        Task ConvertListFormatToTextAsync(_Document document);
        Task DeleteAllHeadersAndFootersAsync(_Document document);
        Task ClearTabStopsAsync(Paragraph paragraph);
    }
}
