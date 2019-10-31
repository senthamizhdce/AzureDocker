
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading.Tasks;
using System;

namespace AzureFunction
{
    public static class Function1
    {
        [FunctionName("LeaveRequestList")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequest req, TraceWriter log)
        {
            List<LeaveRequest> lst = null;
            string EmployeeId = string.Empty;

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            EmployeeId = data?.employeeId;


            if (EmployeeId != null)
            {
                lst = new List<LeaveRequest>();
                for (int i = 1; i < 4; i++)
                {
                    lst.Add(new LeaveRequest() { EmployeeId = "10", Name = $"Employee {i}", From = DateTime.Now.AddDays(-i - 1), To = DateTime.Now.AddDays(-i) });
                }

                return (ActionResult)new OkObjectResult(lst);
            }

            return new BadRequestObjectResult("Please pass a employee Id in the request body");
        }
    }
}
