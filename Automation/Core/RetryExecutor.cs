namespace Automation.Core
{
    using System;
    using System.Threading.Tasks;
    using Argument.Check;
    using Serilog;

    /// <summary>
    /// Retry Executor class.
    /// </summary>
    public class RetryExecutor
    {
        private readonly RetryStrategy _retryStrategy;

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryExecutor"/> class.
        /// </summary>
        /// <param name="retryStrategy">Retry Strategy object.</param>
        public RetryExecutor(RetryStrategy retryStrategy)
        {
            _retryStrategy = retryStrategy;
        }

        /// <inheritdoc/>
        public async Task RetryAsyncTask(Func<Task> action)
        {
            Throw.IfNull(() => action);
            int retries = 0;
            int maxRetries = _retryStrategy.GetMaxRetries();
            TimeSpan interval = _retryStrategy.GetTimeInterval();

            while (true)
            {
                try
                {
                    retries++;
                    await action().ConfigureAwait(false);
                    return;
                }
                catch (Exception ex)
                {
                    Log.Warning(ex.Message);
                    Log.Information($"Retrying for error {ex.Message} for iteration: {retries}");

                    if (retries == maxRetries)
                    {
                        throw new Exception($"{ex.Message} even after {_retryStrategy.GetMaxRetries()} retries", ex);
                    }
                    else
                    {
                        Task.Delay(interval).Wait();
                    }
                }
            }
        }

        /// <inheritdoc/>
        public void Retry(Action action, Action catchBlockAction = null)
        {
            Throw.IfNull(() => action);
            int retries = 0;
            int maxRetries = _retryStrategy.GetMaxRetries();
            TimeSpan interval = _retryStrategy.GetTimeInterval();

            while (true)
            {
                try
                {
                    retries++;
                    action();
                    break;
                }
                catch (Exception ex)
                {
                    Log.Warning(ex.Message);
                    Log.Information($"Retrying for error {ex.Message} for iteration: {retries}");
                    if (retries == maxRetries)
                    {
                        throw new Exception($"{ex.Message} even after {_retryStrategy.GetMaxRetries()} retries", ex);
                    }
                    else
                    {
                        Task.Delay(interval).Wait();
                        catchBlockAction?.Invoke();
                    }
                }
            }
        }
    }
}
