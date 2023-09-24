// See https://benchmarkdotnet.org/articles/guides/how-it-works.html for more information

using BenchmarkDotNet.Running;

// var summary = BenchmarkRunner.Run<CompileVbaProjectBenchmark>();
var summary = BenchmarkRunner.Run<CompileMacroBenchmark>();
