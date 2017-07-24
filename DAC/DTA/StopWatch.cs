using System.Diagnostics;

namespace DAC
{
    class StopWatch
    {
        private static Stopwatch stopWatch = new Stopwatch();
        private static long milliSeconds = 0;

        /// <summary>
        /// NOP(1000000) == wait 1 seconds.
        /// </summary>
        /// <param name="durationTicks">in micro seconds</param>
        public static void NOP(long durationTicks)
        {
            durationTicks = durationTicks * Stopwatch.Frequency / 1000000;
            stopWatch.Start();

            while (stopWatch.ElapsedTicks < durationTicks)
            {
            }

            milliSeconds = stopWatch.ElapsedMilliseconds;
            stopWatch.Stop();
            stopWatch.Reset();
        }
    }
}
