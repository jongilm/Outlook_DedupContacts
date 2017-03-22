
// http://en.csharp-online.net/Delegates_and_Events%E2%80%94Callback_Methods

#region Using directives
 
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
 
#endregion
 
namespace AsynchDelegates
{
   //////////////////////////////////////
   // class ClassWithDelegate
   //////////////////////////////////////
   public class ClassWithDelegate
   {
      // A multicast delegate that encapsulates a method that returns an int
      public delegate int DelegateThatReturnsInt();
      public event DelegateThatReturnsInt theDelegate;

      public void Run()
      {
         for ( ; ; )
         {
            // Sleep for a half second
            Thread.Sleep( 500 );
            //CoWaitForMultipleHandles(COWAIT_WAITALL,50,0,null,0); // Message Pump
            if ( theDelegate != null )
            {
               // Explicitly invoke each delegated method
               foreach (DelegateThatReturnsInt del in theDelegate.GetInvocationList() )
               {
                  // Invoke asynchronously
                  // Pass the delegate in as a state object
                  del.BeginInvoke( new AsyncCallback( ResultsReturned ), del );
               }
            }
         }
      } // Run()
 
      // Call back method to capture results
      private void ResultsReturned( IAsyncResult iar )
      {
         DelegateThatReturnsInt del;
         int result;

         // Cast the state object back to the delegate type
         del = ( DelegateThatReturnsInt ) iar.AsyncState;
 
         // Call EndInvoke on the delegate to get the results
         result = del.EndInvoke( iar );
 
         // Display the results
         Console.WriteLine( "Delegate returned result: {0}", result );
      } // ResultsReturned()

   } // end class ClassWithDelegate
 
   //////////////////////////////////////
   // class FirstSubscriber
   //////////////////////////////////////
   public class FirstSubscriber
   {
      private int myCounter = 0;

      public void Subscribe( ClassWithDelegate theClassWithDelegate )
      {
         theClassWithDelegate.theDelegate += new ClassWithDelegate.DelegateThatReturnsInt( DisplayCounter );
      }
   
      public int DisplayCounter()
      {
         Console.WriteLine( "Busy in DisplayCounter..." );
         Thread.Sleep( 10000 );

         Console.WriteLine( "Done with work in DisplayCounter..." );
         myCounter++;

         return myCounter;
      }
   } // end class FirstSubscriber 
 
   //////////////////////////////////////
   // class SecondSubscriber
   //////////////////////////////////////
   public class SecondSubscriber
   {
      private int myCounter = 0;

      public void Subscribe( ClassWithDelegate theClassWithDelegate )
      {
         theClassWithDelegate.theDelegate += new ClassWithDelegate.DelegateThatReturnsInt( Doubler );
      }
   
      public int Doubler()
      {
         return myCounter += 2;
      }
   } // end class SecondSubscriber
 
   //////////////////////////////////////
   // class Test
   //////////////////////////////////////
   public class Asynch_Test
   {
      public static void Asynch_Main()
      {
         ClassWithDelegate theClassWithDelegate;
         FirstSubscriber fs;
         SecondSubscriber ss;

         theClassWithDelegate = new ClassWithDelegate();

         fs = new FirstSubscriber();
         fs.Subscribe( theClassWithDelegate );

         ss = new SecondSubscriber();
         ss.Subscribe( theClassWithDelegate );

         theClassWithDelegate.Run();
      } // Main()

   } // end class Asynch_Test
}
