using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Grebarz古霸子
{
    public class WooCommerceLicenseManager
    {
        private const string ConsumerKey = "ck_b0b0d70fb94494c7144b3fce3c448fafaa9eba7d";
        private const string ConsumerSecret = "cs_ba0aa8a42ff717ef7fa91f769825f7c49a8b3a4e";

        public async Task<bool> ValidateLicenseKeyAsync(string licenseKey)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var credentials = $"{ConsumerKey}:{ConsumerSecret}";
                    var byteArray = Encoding.ASCII.GetBytes(credentials);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
                    client.Timeout = TimeSpan.FromSeconds(10);

                    var requestUrl = $"https://gz-rebarplanner.com/wp-json/lmfwc/v2/licenses/{licenseKey}?consumer_key=" + ConsumerKey +"?consumer_secret=" + ConsumerSecret;
                    string strSendRequest = "\nSending request to: " + requestUrl; // Debug log
                    strSendRequest.Gz_StatusBarMsg();

                    var response = await client.GetAsync(requestUrl);
                    response.EnsureSuccessStatusCode();

                    var result = await response.Content.ReadAsStringAsync();
                    string strResponse = "\nResponse received: " + result; // Debug log
                    strResponse.Gz_StatusBarMsg();

                    dynamic json = System.Text.Json.JsonDocument.Parse(result).RootElement;
                    if (json.success == true && json.data.status == 2)
                    {
                        DateTime expiresAt = DateTime.Parse(json.data.expiresAt.ToString());
                        return expiresAt > DateTime.Now;
                    }

                    return false;
                }
            }
            catch (HttpRequestException httpEx)
            {
                string str = "\nHTTP Error during license validation: " + httpEx.Message;
                str.Gz_StatusBarMsg();
                return false;
            }
            catch (TaskCanceledException)
            {
                string str = "\nRequest timed out during license validation.";
                str.Gz_StatusBarMsg();
                return false;
            }
            catch (Exception ex)
            {
                string str = "\nError during license validation: " + ex.Message;
                str.Gz_StatusBarMsg();
                return false;
            }
        }
    }
}
