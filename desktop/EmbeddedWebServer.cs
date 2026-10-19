using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace NoppaGovernanceApp
{
    public class EmbeddedWebServer : IDisposable
    {
        private readonly HttpListener _listener;
        private readonly int _port;
        private readonly CancellationTokenSource _cts = new();
        private readonly Dictionary<string, byte[]> _memoryFiles = new(StringComparer.OrdinalIgnoreCase);
        private readonly string? _diskFolder;
        private bool _isRunning;

        public int Port => _port;
        public string BaseUrl => $"http://localhost:{_port}";

        public EmbeddedWebServer(int port = 3334)
        {
            _port = port;
            _listener = new HttpListener();

            // Try to find local 'out' folder first (for development or alongside exe)
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string candidateOut = Path.Combine(appDir, "out");
            if (Directory.Exists(candidateOut))
            {
                _diskFolder = candidateOut;
            }
            else
            {
                string candidateDevOut = Path.GetFullPath(Path.Combine(appDir, "..", "..", "..", "..", "out"));
                if (Directory.Exists(candidateDevOut))
                {
                    _diskFolder = candidateDevOut;
                }
            }

            // Always load embedded zip as base/fallback
            LoadEmbeddedZip();

            // Setup prefixes
            bool prefixAdded = false;
            try
            {
                _listener.Prefixes.Add($"http://localhost:{_port}/");
                prefixAdded = true;
            }
            catch { }

            try
            {
                _listener.Prefixes.Add($"http://127.0.0.1:{_port}/");
                prefixAdded = true;
            }
            catch { }

            if (!prefixAdded)
            {
                throw new InvalidOperationException($"Kon geen HTTP listener prefixes registreren op poort {_port}.");
            }
        }

        private void LoadEmbeddedZip()
        {
            var assembly = Assembly.GetExecutingAssembly();
            // Look for embedded wwwroot.zip
            string resourceName = null!;
            foreach (var name in assembly.GetManifestResourceNames())
            {
                if (name.EndsWith("wwwroot.zip", StringComparison.OrdinalIgnoreCase))
                {
                    resourceName = name;
                    break;
                }
            }

            if (resourceName != null)
            {
                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
                    foreach (var entry in archive.Entries)
                    {
                        if (string.IsNullOrEmpty(entry.Name)) continue; // Directory entry
                        string normPath = entry.FullName.Replace('\\', '/').TrimStart('/');
                        using var entryStream = entry.Open();
                        using var ms = new MemoryStream();
                        entryStream.CopyTo(ms);
                        _memoryFiles[normPath] = ms.ToArray();
                    }
                }
            }
        }

        public void Start()
        {
            if (_isRunning) return;
            _listener.Start();
            _isRunning = true;
            Task.Run(ListenLoop, _cts.Token);
        }

        private async Task ListenLoop()
        {
            while (!_cts.Token.IsCancellationRequested && _listener.IsListening)
            {
                try
                {
                    var context = await _listener.GetContextAsync();
                    _ = Task.Run(() => HandleRequest(context));
                }
                catch (HttpListenerException)
                {
                    break;
                }
                catch (ObjectDisposedException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Listener error: {ex.Message}");
                }
            }
        }

        private void HandleRequest(HttpListenerContext context)
        {
            var request = context.Request;
            var response = context.Response;

            try
            {
                string rawPath = request.Url?.AbsolutePath ?? "/";
                string decodedPath = Uri.UnescapeDataString(rawPath).TrimStart('/');
                if (string.IsNullOrWhiteSpace(decodedPath))
                {
                    decodedPath = "index.html";
                }

                byte[]? fileBytes = null;
                string matchedPath = decodedPath;

                // 1. Try disk if available
                if (!string.IsNullOrEmpty(_diskFolder))
                {
                    string diskPath = Path.Combine(_diskFolder, decodedPath.Replace('/', Path.DirectorySeparatorChar));
                    if (File.Exists(diskPath))
                    {
                        fileBytes = File.ReadAllBytes(diskPath);
                    }
                }

                // 2. Try memory
                if (fileBytes == null)
                {
                    if (_memoryFiles.TryGetValue(decodedPath, out var bytes))
                    {
                        fileBytes = bytes;
                    }
                    else if (_memoryFiles.TryGetValue(decodedPath + ".html", out bytes))
                    {
                        fileBytes = bytes;
                        matchedPath = decodedPath + ".html";
                    }
                    else if (_memoryFiles.TryGetValue(decodedPath + "/index.html", out bytes))
                    {
                        fileBytes = bytes;
                        matchedPath = decodedPath + "/index.html";
                    }
                    else if (!Path.HasExtension(decodedPath) && _memoryFiles.TryGetValue("index.html", out bytes))
                    {
                        // SPA fallback
                        fileBytes = bytes;
                        matchedPath = "index.html";
                    }
                }

                if (fileBytes != null)
                {
                    response.ContentType = GetContentType(matchedPath);
                    response.ContentLength64 = fileBytes.Length;
                    response.StatusCode = 200;
                    response.Headers.Add("Access-Control-Allow-Origin", "*");
                    response.OutputStream.Write(fileBytes, 0, fileBytes.Length);
                }
                else
                {
                    response.StatusCode = 404;
                    byte[] notFound = System.Text.Encoding.UTF8.GetBytes("Bestand niet gevonden");
                    response.OutputStream.Write(notFound, 0, notFound.Length);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    response.StatusCode = 500;
                    byte[] err = System.Text.Encoding.UTF8.GetBytes($"Server fout: {ex.Message}");
                    response.OutputStream.Write(err, 0, err.Length);
                }
                catch { }
            }
            finally
            {
                try
                {
                    response.OutputStream.Close();
                }
                catch { }
            }
        }

        private static string GetContentType(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext switch
            {
                ".html" or ".htm" => "text/html; charset=utf-8",
                ".js" or ".mjs" => "application/javascript; charset=utf-8",
                ".css" => "text/css; charset=utf-8",
                ".json" => "application/json; charset=utf-8",
                ".svg" => "image/svg+xml",
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".ico" => "image/x-icon",
                ".woff" => "font/woff",
                ".woff2" => "font/woff2",
                ".ttf" => "font/ttf",
                ".txt" => "text/plain; charset=utf-8",
                _ => "application/octet-stream"
            };
        }

        public void Dispose()
        {
            if (_isRunning)
            {
                _isRunning = false;
                try { _cts.Cancel(); } catch { }
                try { _listener.Stop(); } catch { }
                try { _listener.Close(); } catch { }
            }
        }
    }
}
