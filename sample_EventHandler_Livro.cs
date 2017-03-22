using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;

namespace Livro
{
  public delegate void AgeChangedEventHandler(object sender, AgeChangedEventArgs e);

  public class AgeChangedEventArgs:EventArgs
  {
    private Int32 _age2;

    internal AgeChangedEventArgs(int _age)
    {
      this._age2 = _age;
    }

    public Int32 Age
    {
      get { return _age2; }
    }
  }

  public class Person
  {
    private Int32 _age1;
    public event AgeChangedEventHandler AgeChanged;

    public Int32 Age
    {
      get
      {
        return _age1;
      }
      set
      {
        if( _age1 == value)
        {
          return;
        }

        _age1 = value;
        OnAgeChanged(new AgeChangedEventArgs(_age1));
      }
    }

    protected virtual void OnAgeChanged(AgeChangedEventArgs e)
    {
      if( AgeChanged != null )
      {
        AgeChanged(this, e);
      }
    }
  }
  
  class Livro_Start
  {
    public static void Livro_Main()
    {
      Person p = new Person();
      p.AgeChanged += PrintAge;
      p.Age = 20;
    }

    public static void PrintAge(object sender, AgeChangedEventArgs e)
    {
      Console.WriteLine(e.Age);
    }
  }
} 
