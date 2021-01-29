Imports NUnit.Framework
Imports PeanutButter.INI
Imports n = NUnit.Framework

Namespace VBLineParser
    <TestFixture>
    Public Class Tests
        <Test>
        public Sub CustomLineParserUsage()
            ' Arrange
            Dim parser = new MyLineParser()
            Dim sut = new INIFile()
            sut.CustomLineParser = parser
            sut.ParseStrategy = ParseStrategies.Custom
            
            ' Act
            sut.Parse("[section]" & vbCrlf & "key=value")
            
            ' Assert
            Assert.That(sut("section")("key"), n.Is.EqualTo("value"))
        End Sub
    End Class
    
    Public Class MyLineParser 
        Implements ILineParser

        Public Function Parse(line As String) As IParsedLine Implements ILineParser.Parse
            if (line.StartsWith("["))
                return new MyParsedLine(line.Trim(), Nothing, "", false)
            End If
            Dim parts = line.Split("=")
            Dim key = parts.First()
            Dim value = StripComments(String.Join("=", parts.Skip(1)))
            
            return new MyParsedLine(key, value, "", false)
        End Function
        
        Shared Quote as Char = """"

        Private Function StripComments(s As String) As String
            if s.Length < 2 Or s(0) <> Quote Or s.Last() <> Quote
                return s
            End If
            return string.Join("", s.Skip(1).Take(s.Length - 2))
        End Function
    End Class
    
    public class MyParsedLine
        Implements IParsedLine
        Public ReadOnly Property Key As String Implements IParsedLine.Key
        Public ReadOnly Property Value As String Implements IParsedLine.Value
        Public ReadOnly Property Comment As String Implements IParsedLine.Comment
        Public ReadOnly Property Escaped As Boolean
        Public ReadOnly Property ContainedEscapedEntities As Boolean Implements IParsedLine.ContainedEscapedEntities
        
        public Sub New(key as String, value as string, comment as string, escaped as Boolean)
            Me.Value = value
            Me.Comment = comment
            Me.Escaped = escaped
            Me.Key = key
        End Sub
    End class
    
End Namespace