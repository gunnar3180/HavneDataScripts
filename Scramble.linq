<Query Kind="Program" />

void Main()
{
	var alfaNum = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789abcdefghijklmnopqrstuvwxyz";
	int maxPos = alfaNum.Length - 1;
	Console.Write("Passord: ");
	string password = Console.ReadLine();
	var coded = new List<char>();
	foreach (var c in password)
	{
		int index = alfaNum.IndexOf(c);
		int newIndex = (index + 52) % maxPos;
		coded.Add(alfaNum[newIndex]);
	}
	
	Console.WriteLine($"Kodet passord = \"{new string(coded.ToArray())}\"");
	
	var decoded = new List<char>();
	foreach (var c in coded)
	{
		int index = alfaNum.IndexOf(c);
		int newIndex = index - 52;
		if (newIndex < 0)
		{
			newIndex += maxPos;
		}
		newIndex = newIndex % maxPos;
		decoded.Add(alfaNum[newIndex]);
	}
	
	//Console.WriteLine($"Dekodet passord = \"{new string(decoded.ToArray())}\"");
}

// Define other methods and classes here