using System.Diagnostics.CodeAnalysis;

namespace GemboxWithRenderIssues.Helpers
{
    // THE CODE BELOW HAS BEEN COPIED FROM RADIUS.REPORTING
    [SuppressMessage("Globalization", "CA1305: Specify IFormatProvider", Justification = "Code copied from Radius.Reporting")]
    public sealed class MergeFieldParameters
    {
        private readonly Dictionary<string, string> _parameters;

        private MergeFieldParameters(Dictionary<string, string> parameters)
        {
            _parameters = parameters;
        }

        public T GetParameter<T>(string key, T defaultValue = default) where T : struct, IConvertible
        {
            if (!_parameters.TryGetValue(key, out var value))
                return defaultValue;

            return (T)Convert.ChangeType(value, typeof(T));
        }

        public static MergeFieldParameters FromFieldName(string fieldName)
        {
            var parameters = "";
            var pos = fieldName.IndexOf("(", StringComparison.Ordinal);
            if (pos >= 0)
                parameters = fieldName.Substring(pos + 1, fieldName.IndexOf(")", pos, StringComparison.Ordinal) - pos - 1);
            if (string.IsNullOrEmpty(parameters))
                return new MergeFieldParameters(new Dictionary<string, string>());
            var dict = parameters
                .Split(',')
                .Select(prm => prm.Split(':'))
                .ToDictionary(prm => prm[0], prm => prm[1]);
            return new MergeFieldParameters(dict);
        }
    }
}