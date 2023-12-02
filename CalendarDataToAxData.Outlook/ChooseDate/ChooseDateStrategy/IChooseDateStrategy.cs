using System;

public interface IChooseDateStrategy
{
    DateTime From { get; }
    DateTime To { get; }
}