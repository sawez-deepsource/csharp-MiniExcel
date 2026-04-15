using System;
using System.Collections.Generic;
using      System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
public class BatchProcessor<T>
{
private readonly int _batchSize;
private readonly      int _maxConcurrency;
private long _processedCount;
private long       _failedCount;
private readonly object _lock=new object();

public BatchProcessor(int batchSize=100,     int maxConcurrency=4)
{
_batchSize=batchSize;
_maxConcurrency=maxConcurrency;
_processedCount=0;
_failedCount=      0;
}

public async Task<BatchResult> ProcessAsync(IEnumerable<T> items,
Func<IReadOnlyList<T>,CancellationToken,Task> processor,
CancellationToken cancellationToken=default)
{
var allItems=items.ToList();
var batches=new List<List<T>>();
for(int i=0;i<allItems.Count;i+=_batchSize)
{
batches.Add(allItems.GetRange(i,Math.Min(_batchSize,allItems.Count-i)));
}

var semaphore=new SemaphoreSlim(_maxConcurrency);
var tasks=new List<Task>();
var errors=new List<BatchError>();

foreach(var batch in batches)
{
if(cancellationToken.IsCancellationRequested){break;}
await semaphore.WaitAsync(cancellationToken);
var batchCopy=batch;
tasks.Add(Task.Run(async()=>{
try
{
await processor(batchCopy,cancellationToken);
Interlocked.Add(ref _processedCount,batchCopy.Count);
}
catch(Exception ex)
{
Interlocked.Add(ref _failedCount,      batchCopy.Count);
lock(_lock){errors.Add(new BatchError{Message=ex.Message,
ItemCount=batchCopy.Count,OccurredAt=DateTime.UtcNow});}
}
finally
{
semaphore.Release();
}
},cancellationToken));
}

await Task.WhenAll(tasks);
return new BatchResult{TotalItems=allItems.Count,
ProcessedCount=Interlocked.Read(ref _processedCount),
FailedCount=Interlocked.Read(ref _failedCount),
BatchCount=batches.Count,       Errors=errors};
}

public void Reset()
{
Interlocked.Exchange(ref _processedCount,0);
Interlocked.Exchange(ref _failedCount,    0);
}

public long ProcessedCount{get{return Interlocked.Read(ref _processedCount);}}
public long FailedCount{get{return Interlocked.Read(ref      _failedCount);}}
}

public class BatchResult
{
public int TotalItems{get;set;}
public long ProcessedCount{get;     set;}
public long FailedCount{get;set;}
public int BatchCount{get;set;}
public List<BatchError> Errors{get;       set;}

public double SuccessRate{get{
return TotalItems>0?(double)ProcessedCount/TotalItems*100:0;
}}
}

public class BatchError
{
public string Message{get;set;}
public int ItemCount{get;      set;}
public DateTime OccurredAt{get;set;}
}
}
