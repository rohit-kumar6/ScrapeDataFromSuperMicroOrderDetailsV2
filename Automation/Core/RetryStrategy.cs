namespace Automation.Core
{
    using System;

    /// <summary>
    /// Retry strategy for retry executions.
    /// </summary>
    public class RetryStrategy
    {
        private readonly int _maxRetries;
        private readonly int _incrementFactor;
        private readonly TimeSpan _interval;
        private readonly TimeSpan _maxInterval;

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryStrategy"/> class.
        /// </summary>
        /// <param name="maxRetries">Number of max retries.</param>
        /// <param name="interval">Time interval between each retry.</param>
        public RetryStrategy(int maxRetries, TimeSpan interval)
        {
            _maxRetries = maxRetries;
            _interval = interval;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="RetryStrategy"/> class.
        /// </summary>
        /// <param name="interval">Initial time interval between each retry.</param>
        /// <param name="incrementFactor">Time interval multipying factor.</param>
        /// <param name="maxInterval">Max time interval between each retry.</param>
        public RetryStrategy(TimeSpan interval, int incrementFactor, TimeSpan maxInterval)
        {
            _interval = interval;
            _incrementFactor = incrementFactor;
            _maxInterval = maxInterval;
        }

        /// <summary>
        /// Get max retries.
        /// </summary>
        /// <returns>Max number of retries.</returns>
        public int GetMaxRetries()
        {
            return _maxRetries;
        }

        /// <summary>
        /// Get increment factor of time interval.
        /// </summary>
        /// <returns>Time interval increment factor.</returns>
        public int GetIncrementFactor()
        {
            return _incrementFactor;
        }

        /// <summary>
        /// Get time interval.
        /// </summary>
        /// <returns>Interval in TimeSpan.</returns>
        public TimeSpan GetTimeInterval()
        {
            return _interval;
        }

        /// <summary>
        /// Get max time interval.
        /// </summary>
        /// <returns>Max time interval in TimeSpan.</returns>
        public TimeSpan GetMaxTimeInterval()
        {
            return _maxInterval;
        }
    }
}
