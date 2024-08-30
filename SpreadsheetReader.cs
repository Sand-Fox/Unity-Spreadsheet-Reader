using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.Networking;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Globalization;
using System;

[Serializable]
public class SpreadsheetReader<T> where T : new()
{
    public const string FormatExport = @"https://docs.google.com/spreadsheets/d/{0}/gviz/tq?tqx=out:csv&sheet={1}";
    public const string FormatDocumentID = @"https://docs.google.com/spreadsheets/d/(?<ID>.*)/";

    [SerializeField] private string _documentURL;
    [SerializeField] private string _sheetName;

    [Space]
    public List<T> Items;

    public IEnumerator ReadSpreadsheet()
    {
        string url = string.Format(FormatExport, GetDocumentID(_documentURL), _sheetName);
        UnityWebRequest request = UnityWebRequest.Get(url);

        yield return request.SendWebRequest();

        if (request.result == UnityWebRequest.Result.ConnectionError || request.result == UnityWebRequest.Result.ProtocolError)
        {
            Debug.LogWarning("Failed to read spreadsheet at url : " + url);
        }
        else
        {
            string textCSV = request.downloadHandler.text;
            ReadLines(textCSV);
        }
        request.Dispose();
    }

    private void ReadLines(string textCSV)
    {
        string[] lines = Regex.Split(textCSV, "\r\n|\r|\n");
        string[] headers = Split(lines[0]);
        Items = new();

        for (int i = 1; i < lines.Length; i++)
        {
            T item = new();
            string[] cells = Split(lines[i]);
            if (string.IsNullOrWhiteSpace(cells[0])) continue;

            ReadCells(item, headers, cells);
            Items.Add(item);
        }
    }

    private void ReadCells(T item, string[] headers, string[] cells)
    {
        Type type = typeof(T);

        for (int i = 0; i < headers.Length; i++)
        {
            string header = headers[i];
            if (string.IsNullOrWhiteSpace(header)) continue;

            string cell = cells[i];
            FieldInfo field = type.GetField(header);

            if (field == null)
            {
                Debug.Log($"Field {header} not found in object of type {type}");
                continue;
            }

            field.SetValue(item, Parse(cell, field.FieldType));
        }
    }

    #region Utilities

    /// <summary>
    /// Get the XXXX part in URL of type : https://docs.google.com/spreadsheets/d/XXXX/edit?gid=0
    /// </summary>
    private string GetDocumentID(string url)
    {
        Match match = Regex.Match(url, FormatDocumentID);

        if (!match.Success)
        {
            Debug.LogWarning("Spreadsheet url is incorrect.");
            return default;
        }

        return match.Groups["ID"].Value;
    }

    /// <summary>
    /// Split a line by looking for passages in quotes, separated by commas.
    /// </summary>
    private string[] Split(string line)
    {
        bool isInsideQuotes = false;
        List<string> results = new();
        string temp = string.Empty;

        foreach (char character in line)
        {
            if (character == '"')
            {
                isInsideQuotes = !isInsideQuotes;
                continue;
            }

            if (character == ',' && !isInsideQuotes)
            {
                results.Add(temp);
                temp = string.Empty;
                continue;
            }

            temp += character;
        }

        if (temp != string.Empty)
        {
            results.Add(temp);
        }

        return results.ToArray();
    }

    /// <summary>
    /// Parse a string to the given type.
    /// </summary>
    private object Parse(string s, Type type)
    {
        if (type == typeof(string)) return s;

        if (type == typeof(int))
        {
            if (int.TryParse(s, out var result)) return result;
        }

        if (type == typeof(byte))
        {
            if (byte.TryParse(s, out var result)) return result;
        }

        if (type == typeof(short))
        {
            if (short.TryParse(s, out var result)) return result;
        }

        if (type == typeof(long))
        {
            if (long.TryParse(s, out var result)) return result;
        }

        if (type == typeof(bool))
        {
            if (bool.TryParse(s, out var result)) return result;
        }

        if (type == typeof(float))
        {
            if (float.TryParse(s.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var resultFloat))
                return resultFloat;
        }

        if (type == typeof(double))
        {
            if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var resultFloat))
                return resultFloat;
        }

        if (type.IsEnum)
        {
            object result;

            try
            {
                result = Enum.Parse(type, s, true);
            }
            catch (ArgumentException)
            {
                result = default;
            }
            return result;
        }

        return default;
    }

    #endregion
}
