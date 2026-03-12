using System;

namespace APIKeyGenerator
{
    public class Class1
    {
        public void KeyGen()
        {

            var client = new RestClient("https://makevalegroup.scinote.net/oauth/token?grant_type=refresh_token&client_id=d5843a09-f86a-4526-a484-1de8602a02de&client_secret=2a794653-01fa-4acb-a473-1c3b1d9fbf55&refresh_token=YirjHwcE5gKrnFSOfsx6uN2U1PSD_SyqZhYiiCANAkQ&redirect_uri=urn:ietf:wg:oauth:2.0:oob");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            IRestResponse response = client.Execute(request);
            Console.WriteLine(response.Content);
        }

    }
}
