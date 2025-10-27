// See https://benchmarkdotnet.org/articles/guides/how-it-works.html for more information

using BenchmarkDotNet.Running;

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
