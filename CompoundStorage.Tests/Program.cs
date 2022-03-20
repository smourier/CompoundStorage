using System;
using System.IO;
using CompoundStorage;
using CompoundStorage.Utilities;

namespace MyApp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Windows\Recent\AutomaticDestinations", @"a52b0784bd667468.automaticDestinations-ms");
            //path = @"This is a PDF file.doc";

            using (var pset = new Storage(path))
            {
                foreach (var element in pset.Elements)
                {
                    if (element is StreamElement streamElement)
                    {
                        using (var stream = streamElement.OpenStream())
                        {
                            var link = Link.Load(stream, false);
                            if (link != null)
                            {
                                Console.WriteLine(link);
                            }
                            else
                            {
                                Console.WriteLine(element);
                                using (var file = File.OpenWrite("stream" + element.Name))
                                {
                                    stream.CopyTo(file);
                                }
                            }
                        }
                    }
                }

                foreach (var storage in pset.PropertyStorages)
                {
                    Console.WriteLine(storage);
                    foreach (var prop in storage.Properties)
                    {
                        Console.WriteLine(" " + prop);
                    }
                }
            }
        }
    }
}