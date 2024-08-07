﻿namespace RxBim.Tools.TableBuilder.Excel.Tests;

using Di;

public abstract class TestsBase
{
    public TestsBase()
    {
        Container = new DiContainer();
        Container.AddExcelTableBuilder();
    }

    public IContainer Container { get; }
}