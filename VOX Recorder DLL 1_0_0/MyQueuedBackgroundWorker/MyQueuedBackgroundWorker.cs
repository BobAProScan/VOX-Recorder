using System;
using System.Collections.Generic;
using System.ComponentModel;

//https://stackoverflow.com/questions/15119974/creating-backgroundworker-with-queue
public static class MyQueuedBackgroundWorker
{
    public static void QueueWorkItem<Tin, Tout>(
        Queue<QueueItem<Tin>> queue,
        Tin inputArgument,
        Func<DoWorkArgument<Tin>, Tout> doWork,
        Action<WorkerResult<Tout>> workerCompleted)
    {
        if (queue == null) throw new ArgumentNullException("queue");

        BackgroundWorker bw = new BackgroundWorker();
        bw.WorkerReportsProgress = false;
        bw.WorkerSupportsCancellation = false;
        bw.DoWork += (sender, args) =>
        {
            if (doWork != null)
            {
                args.Result = doWork(new DoWorkArgument<Tin>((Tin)args.Argument));
            }
        };
        bw.RunWorkerCompleted += (sender, args) =>
        {
            if (workerCompleted != null)
            {
                workerCompleted(new WorkerResult<Tout>((Tout)args.Result, args.Error));
            }
            queue.Dequeue();
            if (queue.Count > 0)
            {
                QueueItem<Tin> nextItem = queue.Peek(); // as QueueItem<T>;
                nextItem.BackgroundWorker.RunWorkerAsync(nextItem.Argument);
            }
        };

        queue.Enqueue(new QueueItem<Tin>(bw, inputArgument));
        if (queue.Count == 1)
        {
            QueueItem<Tin> nextItem = queue.Peek() as QueueItem<Tin>;
            nextItem.BackgroundWorker.RunWorkerAsync(nextItem.Argument);
        }
    }
}

public class DoWorkArgument<T>
{
    public DoWorkArgument(T argument)
    {
        this.Argument = argument;
    }

    public T Argument { get; private set; }
}

public class WorkerResult<T>
{
    public WorkerResult(T result, Exception error)
    {
        this.Result = result;
        this.Error = error;
    }

    public T Result { get; private set; }

    public Exception Error { get; private set; }
}

public class QueueItem<T>
{
    public QueueItem(BackgroundWorker backgroundWorker, T argument)
    {
        this.BackgroundWorker = backgroundWorker;
        this.Argument = argument;
    }

    public T Argument { get; private set; }

    public BackgroundWorker BackgroundWorker { get; private set; }
}