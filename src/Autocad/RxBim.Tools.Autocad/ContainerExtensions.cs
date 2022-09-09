﻿namespace RxBim.Tools.Autocad
{
    using Di;

    /// <summary>
    /// Extensions for <see cref="IContainer"/>.
    /// </summary>
    public static class ContainerExtensions
    {
        /// <summary>
        /// Adds AutoCAD services to the container.
        /// </summary>
        /// <param name="container"><see cref="IContainer"/> object.</param>
        public static IContainer AddAutocadTools(this IContainer container)
        {
            container.AddSingleton<IDocumentService, DocumentService>();
            container.AddSingleton<IObjectsSelectionService, ObjectsSelectionService>();
            container.AddSingleton<ICommandLineService, CommandLineService>();
            container.AddSingleton<IElementsDisplay, ElementsDisplayService>();
            container.AddSingleton<IProblemElementsStorage, ProblemElementsStorage>();
            container.AddSingleton<ITransactionContextService<DatabaseContext>, DatabaseContextService>();
            container.AddSingleton<ITransactionContextService<DocumentContext>, DocumentContextService>();
            container.AddTransactionServices<AutocadTransactionFactory>();
            return container;
        }
    }
}