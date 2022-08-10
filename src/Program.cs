
using Azure.Identity;
using Dasync.Collections;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph;
using System.Diagnostics;

using IHost host = Host.CreateDefaultBuilder(args).Build();
IConfiguration config = host.Services.GetRequiredService<IConfiguration>();

/*
 * Configuration settings:
 *  messageCount - number of messages to retrieve from mailbox
 *  maxDegreeOfParallelism - number of messages to process in parallel
 *  tenantId - id of AAD tenant
 *  clientId - id of client application configured in tenant. No specific permissions are required 
 */

var messageCount = config.GetValue<int?>("messageCount") ?? 20;
var maxDegreeOfParallelism = config.GetValue<int?>("maxDegreeOfParallelism") ?? 10;
var tenantId = config.GetValue<string>("tenantId");
var clientId = config.GetValue<string>("clientId");

var scopes = new[] { "User.Read", "Mail.Read" };

// uses the device code flow - https://docs.microsoft.com/dotnet/api/azure.identity.devicecodecredential

var options = new TokenCredentialOptions
{
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
};

Func<DeviceCodeInfo, CancellationToken, Task> callback = (code, cancellation) =>
{
    Console.WriteLine(code.Message);
    return Task.FromResult(0);
};

var deviceCodeCredential = new DeviceCodeCredential(
    callback, tenantId, clientId, options);

var graphClient = new GraphServiceClient(deviceCodeCredential, scopes);

//****************************************************************************

Stopwatch stopWatchMain = new Stopwatch();

var messages = await graphClient.Me.Messages
    .Request()
    .Top(messageCount)
    .Filter("hasAttachments eq true")
    .Select("sender,subject")
    .GetAsync();

Console.WriteLine($"### Starting serial processing - {messageCount} messages");

stopWatchMain.Start();

Stopwatch stopWatchAttachment = new Stopwatch();

var totalSizeAttachments = 0;

foreach (var message in messages)
{
    stopWatchAttachment.Restart();

    var attachments = await graphClient.Me.Messages[message.Id].Attachments
    .Request()
    .GetAsync();

    stopWatchAttachment.Stop();

    var totalSizeAttachmentsForMessage = attachments.Sum(x => x.Size);

    totalSizeAttachments += totalSizeAttachmentsForMessage ?? 0;

    Console.WriteLine($"{message.Id.Substring(message.Id.Length - 12)}, attachments:{attachments.Count}, bytes:{totalSizeAttachmentsForMessage}, time:{stopWatchAttachment.Elapsed.TotalSeconds}");
}

stopWatchMain.Stop();

Console.WriteLine($"### Total time serial: {stopWatchMain.Elapsed.TotalSeconds}, total bytes serial: {totalSizeAttachments}");

//****************************************************************************

messages = await graphClient.Me.Messages
    .Request()
    .Top(messageCount)
    .Filter("hasAttachments eq true")
    .Select("sender,subject")
    .GetAsync();

Console.WriteLine($"### Starting parallel processing - {messageCount} messages, {maxDegreeOfParallelism} in parallel");

stopWatchMain.Restart();

totalSizeAttachments = 0;

await messages.ParallelForEachAsync(async message =>
{

    var stopWatch = new Stopwatch();
    stopWatch.Restart();

    var attachments = await graphClient.Me.Messages[message.Id].Attachments
    .Request()
    .GetAsync();

    stopWatch.Stop();

    var totalSizeAttachmentsForMessage = attachments.Sum(x => x.Size);

    totalSizeAttachments += totalSizeAttachmentsForMessage ?? 0;

    Console.WriteLine($"{message.Id.Substring(message.Id.Length - 12)}, attachments:{attachments.Count}, bytes:{totalSizeAttachmentsForMessage}, time:{stopWatch.Elapsed.TotalSeconds}");

}, maxDegreeOfParallelism: maxDegreeOfParallelism);

stopWatchMain.Stop();

Console.WriteLine($"### Total time parallel: {stopWatchMain.Elapsed.TotalSeconds}, total bytes parallel: {totalSizeAttachments}");
