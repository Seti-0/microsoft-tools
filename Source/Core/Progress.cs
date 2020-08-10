using Red.Core.Logs;
using System;
using System.Collections.Generic;
using System.Text;

namespace Red.Core
{
    public static class Progress
    {
        private static float lastProgress = 0;

        private static Action<string> outputAction;

        public static void Init(Action<string> output)
        {
            outputAction = output;
        }

        public static void Reset()
        {
            lastProgress = 0;
        }

        public static void Report(float value)
        {
            if (value - lastProgress >= 0.1)
            {
                lastProgress = value;
                outputAction($"{Math.Round(value * 100)}% complete");
            }
        }

        public static void Complete()
        {
            if (lastProgress != 1)
                outputAction($"100% complete");
        }
    }
}
