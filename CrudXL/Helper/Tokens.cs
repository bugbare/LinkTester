using System;
using System.Collections;

//Declare the tokens class:
public class Tokens : IEnumerable
{
    private string[] elements;

    //define the token constructor method
    Tokens(string source, char[] delimiters)
    {
        // Parse the string into tokens
        elements = source.Split(delimiters);
    }

    /* IEnumerable Interface Implementation, so that a for each statement can also 
     * enumerate the elements of a collection, provided that the collection class 
     * has implemented the System.Collections.IEnumerator and 
     * System.Collections.IEnumerable interfaces*/
     
     /* Declare GetEnumerator() method that is required by IEnumerable. 
      * Returns an instance of a new IEnumerator interface object 
      * implemented through a TokenEnumerator instance
      */
    public IEnumerator GetEnumerator()
    {
        return new TokenEnumerator(this);
    }

    // Inner Class implements IEnumerator interface:
    private class TokenEnumerator : IEnumerator
    {
        private int position = -1;
        private Tokens t;

        public TokenEnumerator(Tokens t)
        {
            this.t = t;
        }

        // Declare the MoveNext method required by IEnumerator:
        public bool MoveNext()
        {
            if (position < t.elements.Length - 1)
            {
                position++;
                return true;
            }
            else
            {
                return false;
            }
        }

        // Declare the Rest method required by IEnumerator:
        public void Reset()
        {
            position = -1;
        }

        // Declare the Current property required by IEnumerator:
        public object Current
        {
            get
            {
                return t.elements[position];
            }
        }
    }

    static void Main()
    {
        // Testing Tokens by breaking the string into tokens:
        Tokens f = new Tokens("This is a well-done program.",
           new char[] { ' ', '-' });
        foreach (string item in f)
        {
            Console.WriteLine(item);
        }
    }
}
