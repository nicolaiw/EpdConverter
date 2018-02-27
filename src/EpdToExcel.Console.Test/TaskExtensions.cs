using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EpdToExcel.Console.Test
{
    public static class TaskExtensions
    {
        // http://blog.danskingdom.com/tag/c-task-thread-throttle-limit-maximum-simultaneous-concurrent-parallel/
        public static async Task<IEnumerable<T>> ForEachAsyncThrottled<T>(this IEnumerable<Task<T>> tasksToRun, int maxConcurrency)
        {
            // Convert to a list of tasks so that we don't enumerate over it multiple times needlessly.
            var tasks = tasksToRun.ToList();

            using (var throttler = new SemaphoreSlim(maxConcurrency))
            {
                var postTaskTasks = new List<Task>();

                // Have each task notify the throttler when it completes so that it decrements the number of tasks currently running.
                tasks.ForEach(t => postTaskTasks.Add(t.ContinueWith(tsk => throttler.Release())));

                // Start running each task.
                foreach (var task in tasks)
                {
                    await throttler.WaitAsync();

                    task.Start();
                }

                // Wait for all of the provided tasks to complete.
                // We wait on the list of "post" tasks instead of the original tasks,
                // otherwise there is a potential race condition where the throttler's using block is exited before some Tasks have had their "post"
                // action completed, which references the throttler, resulting in an exception due to accessing a disposed object.
                var res = await Task.WhenAll(tasks);
                await Task.WhenAll(postTaskTasks);

                return res;
            }
        }
    }
}
