using System;

namespace Com1CMultiRegister {
    public class Program {
        static void Main(string[] args) {
            Manager manager;

            Console.WriteLine("Write path or stay this line empty: ");
            var path = Console.ReadLine();

            Console.WriteLine("Write pattern or stay this line empty: ");
            var pattern = Console.ReadLine();

            if (path == null || path == "") manager = new Manager(pattern);
            else manager = new Manager(path, pattern);

            Console.WriteLine("Press any button to continue.");
            Console.ReadLine();

            var components = manager.GetComponents();
            manager.CreateComponents(components);
        }
    }
}
