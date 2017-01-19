# StringBuilder for VBA

## How to use?

1. Import `src/VBA.StringBuilder.xlsm/StringBuilder.cls` into your VBA project (xlsm, et al).
2. In your VBA IDE, you can instanciate and use StringBuilder class like this:

```vb
Dim sb As New StringBuilder

sb.Append("Hello")
sb.Append(", ")
sb.Append("world")
sb.Append("!")

MsgBox sb.ToString()
```

## Special Thanks

`src/VBA.StringBuilder.xlsm/StringBuilderTest.cls` tests this library with `Assert.bas` in [vbaidiot/Ariawase](https://github.com/vbaidiot/Ariawase).

## License

The Apache License Version 2.0, see LICENSE.txt.