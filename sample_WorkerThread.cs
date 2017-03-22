// http://www.codeproject.com/KB/cs/workerthread.aspx
// MainForm.cs


namespace WorkerThread
{
  // Delegates used to call MainForm functions from worker thread
  public delegate void DelegateAddString(String s);
  public delegate void DelegateThreadFinished();

  public class MainForm : System.Windows.Forms.Form
  {
    // ...

    // worker thread
    Thread m_WorkerThread;

    // Events used to stop worker thread
    ManualResetEvent m_EventStopThread;
    ManualResetEvent m_EventThreadStopped;

    // Delegate instances used to call user interface functions 
    // from worker thread:
    public DelegateAddString m_DelegateAddString;
    public DelegateThreadFinished m_DelegateThreadFinished;

    // ...

    public MainForm()
    {
      InitializeComponent();

      // Initialize delegates
      m_DelegateAddString = new DelegateAddString(this.AddString);
      m_DelegateThreadFinished = new DelegateThreadFinished(this.ThreadFinished);

      // Initialize events
      m_EventStopThread = new ManualResetEvent(false);
      m_EventThreadStopped = new ManualResetEvent(false);
    }

    // ...

    // Start thread button is pressed
    private void btnStartThread_Click(object sender, System.EventArgs e)
    {
      // ...

      // Reset events
      m_EventStopThread.Reset();
      m_EventThreadStopped.Reset();

      // Create worker thread instance
      m_WorkerThread = new Thread(new ThreadStart(this.WorkerThreadFunction));
      m_WorkerThread.Name = "Worker Thread Sample";   // looks nice in Output window
      m_WorkerThread.Start();
    }

    // Worker thread function.
    // Called indirectly from btnStartThread_Click
    private void WorkerThreadFunction()
    {
      LongProcess longProcess;
      longProcess = new LongProcess(m_EventStopThread, m_EventThreadStopped, this);
      longProcess.Run();
    }

    // Stop worker thread if it is running.
    // Called when user presses Stop button or form is closed.
    private void StopThread()
    {
      if ( m_WorkerThread != null  &&  m_WorkerThread.IsAlive )  // thread is active
      {
        // Set event "Stop"
        m_EventStopThread.Set();

        // Wait when thread  will stop or finish
        while (m_WorkerThread.IsAlive)
        {
          // We cannot use here infinite wait because our thread
          // makes syncronous calls to main form, this will cause deadlock.
          // Instead of this we wait for event some appropriate time
          // (and by the way give time to worker thread) and
          // process events. These events may contain Invoke calls.

          if ( WaitHandle.WaitAll((new ManualResetEvent[] {m_EventThreadStopped}), 100, true) )
          {
            break;
          }
          Application.DoEvents();
        }
      }
    }

    // Add string to list box.
    // Called from worker thread using delegate and Control.Invoke

    private void AddString(String s)
    {
      listBox1.Items.Add(s);
    }

    // Set initial state of controls.
    // Called from worker thread using delegate and Control.Invoke
    private void ThreadFinished()
    {
      btnStartThread.Enabled = true;
      btnStopThread.Enabled = false;
    }
  }
}

// LongProcess.cs
namespace WorkerThread
{
  public class LongProcess
  {
    // ...

    // Function runs in worker thread and emulates long process.
    public void Run()
    {
      int i;
      String s;

      for (i = 1; i <= 10; i++)
      {
        // Make step
        s = "Step number " + i.ToString() + " executed";

        Thread.Sleep(400);

        // Make synchronous call to main form.
        // MainForm.AddString function runs in main thread.
        // (To make asynchronous call use BeginInvoke)
        m_form.Invoke(m_form.m_DelegateAddString, new Object[] {s});

        // check if thread is cancelled
        if ( m_EventStop.WaitOne(0, true) )
        {
          // Clean-up operations may be placed here
          // ...

          // Inform main thread that this thread stopped
          m_EventStopped.Set();
          return;
        }
      }

      // Make synchronous call to main form
      // to inform it that thread finished
      m_form.Invoke(m_form.m_DelegateThreadFinished, null);
    }
  }
}
