namespace Excel.Report.PDF
{
    public interface IPostProcessCommand
    {
        void Execute();
    }

    public static class PostProcessCommandExtensions
    {
        public static void ExecuteAll(this IEnumerable<IPostProcessCommand> commands)
        {
            foreach (var command in commands)
            {
                command.Execute();
            }
        }
    }
}
