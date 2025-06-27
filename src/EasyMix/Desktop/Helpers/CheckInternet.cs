using System.Net.Http;

namespace Desktop.Helpers
{
    public class CheckInternet
    {
        public static async Task<bool> IsInternetAvailableAsync()
        {
            try
            {
                using var client = new HttpClient
                {
                    Timeout = TimeSpan.FromSeconds(3)
                };

                using var response = await client.GetAsync("http://www.google.com");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }
    }
}
