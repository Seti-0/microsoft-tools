using Red.Core.Logs;
using System;
using System.Collections.Generic;
using System.Net;
using System.Text;

namespace Red.Core
{
    public static class Flow
    {
        public enum Reason
        {
            Cancel,
            Quit
        }

        public enum State
        {
            Initializing,
            Idle,
            Running,
            Interrupted
        }

        public static State CurrentState { get; private set; } = State.Initializing;

        public static string CurrentActionName { get; private set; } = null;

        public static Reason InterruptReason { get; private set; }

        public static bool Interrupted { get => CurrentState == State.Interrupted; }

        public static bool Idle { get => CurrentState == State.Idle; }

        public static bool Initialized { get => CurrentState != State.Initializing; }

        public static void Init()
        {
            CurrentState = State.Idle;
        }

        public static void Interrupt(Reason reason)
        {
            CurrentState = State.Interrupted;
            InterruptReason = reason;

            Log.Core.Debug("Interrupt signal sent");
        }

        public static void RunWithCancel(string name, Action action, string cancelMessage)
        {
            if (!Idle)
            {
                Log.Core.Error($"Internal Error: Cannot start action: \"{name}\"");
                Log.Core.PushIndent();

                string statusMessage;

                switch (CurrentState)
                {
                    case State.Initializing: statusMessage = "App has not finished initializing"; break;
                    case State.Running: statusMessage = $"Previous action: \"{CurrentActionName}\" has not ended"; break;
                    case State.Interrupted: statusMessage = $"Previous action: \"{CurrentActionName}\" is interrupted, but has not ended"; break;
                    default: statusMessage = "(Unrecognised internal state)"; break;
                }

                Log.Core.Debug(statusMessage);
                Log.Core.Debug("Note to self: See \"Flow.cs\" for further info");
                Log.Core.PopIndent();
                return;
            }

            CurrentState = State.Running;
            CurrentActionName = name;

            action();

            if (Interrupted && InterruptReason == Reason.Cancel)
                Log.Core.Info(cancelMessage);

            if (!Interrupted || InterruptReason == Reason.Cancel)
            {
                CurrentActionName = null;
                CurrentState = State.Idle;
            }
        }
    }
}
