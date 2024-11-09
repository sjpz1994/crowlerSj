using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace crowlerSj.service
{
    public class GoogleSearchService
    {
        private readonly HttpClient _httpClient;

        public GoogleSearchService(HttpClient httpClient)
        {
            _httpClient = httpClient;
        }

        public async Task<string> SearchAndGetResults(string query)
        {
            try
            {
                HttpResponseMessage response = await _httpClient.GetAsync($"https://www.google.com/search?q={query}");
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();

                return responseBody;
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"An error occurred: {e.Message}");
                return null;
            }
        }
    }
}
