using Microsoft.Graph;
using System;
using System.Collections.Generic;

namespace MicrosoftGraphSDK
{
    class Program
    {
        static void Main(string[] args)
        {
            //https://developer.microsoft.com/en-us/graph
            //Navigate to Graph Explorer

            var client = new GraphServiceClient(new DelegateAuthenticationProvider(async request =>
            {
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("bearer", "EwCQA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAVqWx8RzTkHlb0SlveLACO3IvLCcvWMfQRnyxwnFJ26A3HfuwvafD0Hj5ad1E1agIVv//2k0aB7o8eNB/KIZrYa5rMp25GeL5o1O9HIw6LmN4lvss95w51TinG8XsLsp2pmq9R+CLpDqkCQSmytVP8f93am3YsjiN3qPF5Cv06ml6BJPGdEGuTtRFbLikFwsVGMqXx5+CpZu8j1xRpHjvOfNteUJsDS9IhOjD1wnOSBry5VC4s3EBqpBfMHYjfGqqrjUP8COdFfLGGYQ3j4UpB42GqAihbfq9ALPqLLeHUitLn+9gwAODE4QhwPIrfqgDFC7UNjfI/m1WJFJdtj/XlgDZgAACJzejL3lFJ9bYAL8uUqcNLqFaB5rtwvbTl4drtPBB5Pdn/fsF2UcYPm2yW004+73MFaQ1c1zpjj8ZIplwdMT9pP+5IkTPJQXe5TdZLDbZ4GSE7SZDQQeXa6TuTgQ0824xR1ykE9Jnk/UzvYE2cgoGsU6pv6mHSPsAxCi1jFgQl/MHOy+nOinHAQTyE1Meps19ujPBok/O0j/2TsNb0aT3SCZ7Xwe8S9vab+Qn57jGC5KWRMvLcAFuTihcQIrHv7K5qUCjqjGpZbWtVWdFHpa7jnlTjOda7rmDfrJzeVgJXf8EAQ9umAeObyyCDl33rB5jYlaDwGsNiB1Nlnhbtzqv1K5WcFuDCl5EWhObxpGXVIZ1PrSaOJMvQNoAaVsGhBosg7+sDhPxxhJL3/OpSHgKVHMBjJrkucg3t9Y6DOnFvC7CFPoCv5k+1bYcRmw+SeVuMWeLKrDM1+T2L308CabMjBdmd6YtUENbQsXf7lrLawG2YwD5FdxFKnJzJa+Y2V8CfpFM/hjnxxibk6HGcS1D1SM95x/lXgJx74IqIM+rBxmYgM39W/xE2P8XJtxOkuMkRYOdsnG2uVFerIRstB8r7/rExsx05ftPcCtQ7lfzBejH0z8aE0tchriqgFKWVRNK4wGpOVaOe4zYKCQeUabSWnVVhBiKX5gc5+gcFXP4u3soN1ot339dB4ulCsThFYqKemKDCHznb00zaRfAgvMPRzGcA17UlFukYy5F3X+NG0FeY+kQmNgszO+5XhCorYmG3ImQl+gzb5AhQjmGAiGeY1l5akmTBz/FKY4GbYK9Sh93CgQGBHeejVq7KoC");
            }));

            var sendMessage = new Message()
            {
                ToRecipients = new List<Recipient>
                {
                    new Recipient()
                    {
                        EmailAddress=new EmailAddress(){Address="senthamizh.lcs@gmail.com"}
                    }
                },
                Subject = "Hi",
                Body = new ItemBody()
                {
                    ContentType = BodyType.Text,
                    Content = "Body"
                }
            };

            //Bilgates working on it
            //client.Me.SendMail(sendMessage).Request().PostAsync();

            var sendRequest = client.Me.SendMail(sendMessage).Request().GetHttpRequestMessage();

            //var httpClient = new HttpClient();
            var httpClient = GraphClientFactory.CreateClient();
            var result = httpClient.SendAsync(sendRequest);
            result.Wait();

            var message = client.Me.Messages.Request()
                .Select(m => new { m.Subject, m.SentDateTime })
                //.Filter("hasAttachment eq true")
                .Expand(m => m.Attachments)
                .Top(10)
                .GetAsync().Result;


            foreach (var item in message)
            {
                Console.WriteLine(item.Subject + "  -->  " + item.SentDateTime);
            }


            Console.WriteLine("Completed");
            Console.Read();
        }
    }
}
