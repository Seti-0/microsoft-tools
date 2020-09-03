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

                string statusMessage = CurrentState switch
                {
                    State.Initializing => "App has not finished initializing",
                    State.Running => $"Previous action: \"{CurrentActionName}\" has not ended",
                    State.Interrupted => $"Previous action: \"{CurrentActionName}\" is interrupted, but has not ended",
                    _ => "(Unrecognised internal state)",
                };
                
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
