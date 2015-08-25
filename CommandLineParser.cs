using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace bulkmail
{
    public class CommandLineParser
    {
        public static T ParseCommandLine<T>(string[] args) where T : new()
        {
            StringBuilder usage = new StringBuilder();
            try
            {
                List<PropertyInfo> positional = new List<PropertyInfo>();
                Dictionary<string, PropertyInfo> named = new Dictionary<string, PropertyInfo>();
                Dictionary<PropertyInfo, string> required = new Dictionary<PropertyInfo, string>();
                Dictionary<PropertyInfo, string> defaults = new Dictionary<PropertyInfo, string>();

                var tt = typeof(T);
                foreach (var pi in tt.GetProperties())
                {
                    string name = null;
                    {
                        var attr = pi.GetCustomAttribute<PositionalArgAttribute>();
                        if (attr != null)
                        {
                            while (positional.Count <= attr.Position)
                                positional.Add(null);
                            positional[attr.Position] = pi;
                            name = "position " + (attr.Position + 1).ToString();
                        }
                    }
                    if (name == null)
                    {
                        var attr = pi.GetCustomAttribute<NamedArgAttribute>();
                        if (attr != null)
                        {
                            named[attr.ShortForm.ToString().ToLowerInvariant()] = pi;
                            named[attr.LongForm.ToString().ToLowerInvariant()] = pi;
                            name = "-" + attr.ShortForm + "|--" + attr.LongForm;
                        }
                    }
                    bool req = false;
                    {
                        var attr = pi.GetCustomAttribute<RequiredAttribute>();
                        if (attr != null) { required[pi] = name; req = true; }
                    }
                    string def = "";
                    {
                        var attr = pi.GetCustomAttribute<DefaultValueAttribute>();
                        if (attr != null) { def = " = " + (defaults[pi] = attr.Value); }
                    }
                    usage.AppendFormat("  {0}{1}{2}", name, def, (req ? " [required]" : ""));
                    usage.AppendLine();
                    {
                        var attr = pi.GetCustomAttribute<HelpAttribute>();
                        if (attr != null)
                        {
                            usage.AppendFormat("    {0}", attr.Desctiption);
                            usage.AppendLine();
                        }
                    }
                    usage.AppendLine();
                }

                var t = new T();

                Action<PropertyInfo, string> setField = (pi, s) =>
                {
                    if (required.ContainsKey(pi)) required.Remove(pi);
                    if (defaults.ContainsKey(pi)) defaults.Remove(pi);
                    var type = pi.PropertyType;
                    if (type == typeof(string))
                    {
                        pi.SetValue(t, s);
                    }
                    else if (type == typeof(bool))
                    {
                        pi.SetValue(t, true);
                    }
                    else
                    {
                        var p = type.GetMethod("Parse");
                        if (p == null) throw new Exception("the parser cannot handle arguments of type " + type.ToString());
                        try
                        {
                            pi.SetValue(t, p.Invoke(null, new object[] { s }));
                        }
                        catch (Exception exn)
                        {
                            throw new Exception(string.Format("failed to parse {0} as a {1}", s, type.ToString()), exn);
                        }
                    }
                };

                bool dashOptionsAllowed = true;

                int i = 0;
                int pos = 0;
                while (i < args.Length)
                {
                    if (args[i] == "--")
                    {
                        dashOptionsAllowed = false;
                        ++i;
                        continue;
                    }

                    if (!dashOptionsAllowed || !args[i].StartsWith("-", StringComparison.Ordinal))
                    {
                        if (pos >= positional.Count) throw new Exception("too many arguments");
                        setField(positional[pos], args[i]);
                        ++i;
                        ++pos;
                        continue;
                    }

                    string name;
                    if (args[i].StartsWith("--", StringComparison.Ordinal))
                        name = args[i].Substring(2);
                    else
                        name = args[i].Substring(1);
                    name = name.ToLowerInvariant();

                    PropertyInfo pi;
                    if (!named.TryGetValue(name, out pi))
                        throw new Exception("unknown argument " + name);
                    if (pi.PropertyType != typeof(bool)) ++i;
                    setField(pi, args[i]);
                    ++i;
                }

                foreach (var kvp in defaults.ToList())
                    setField(kvp.Key, kvp.Value);

                if (required.Count > 0)
                    throw new Exception("Missing values for required arguments " + string.Join(", ", required.Values));

                return t;
            }
            catch (Exception exn)
            {
                throw new CommandLineParseError("Failed to parse command line: " + exn.Message, usage.ToString(), exn);
            }
        }
    }

    public class CommandLineParseError : Exception
    {
        public CommandLineParseError(string desc, string usage, Exception exn) : base(desc, exn)
        {
            Desc = desc;
            Usage = usage;
        }

        public string Desc;
        public string Usage;
    }

    public class PositionalArgAttribute : Attribute
    {
        public PositionalArgAttribute(int position) { Position = position; }
        public int Position;
    }

    public class NamedArgAttribute : Attribute
    {
        public char ShortForm;
        public string LongForm;
    }

    public class HelpAttribute : Attribute
    {
        public HelpAttribute(string desc) { Desctiption = desc; }
        public string Desctiption;
    }

    public class DefaultValueAttribute : Attribute
    {
        public DefaultValueAttribute(string value) { Value = value; }
        public string Value;
    }

    public class RequiredAttribute : Attribute
    {
    }
}
