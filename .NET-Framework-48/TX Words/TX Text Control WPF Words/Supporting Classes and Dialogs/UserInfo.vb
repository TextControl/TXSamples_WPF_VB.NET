'-----------------------------------------------------------------------------------------------------------
' UserInfo.vb File
'
' Description:
'     Provides information about a user.
'
' copyright:		© Text Control GmbH
'-----------------------------------------------------------------------------------------------------------
Imports System.Security.Cryptography
Imports System.Text

Namespace TXTextControl.Words
    '-----------------------------------------------------------------------------------------------------------
    ' Class UserInfo
    ' Capsulates the user's info like name and password. The password is stored as a SHA1 Hash.
    '-----------------------------------------------------------------------------------------------------------
    Public Class UserInfo

        '-----------------------------------------------------------------------------------------------------------
        ' M E M B E R S
        '-----------------------------------------------------------------------------------------------------------

        Private m_strName As String                   ' User's name
        Private m_rbPasswordHash As Byte()            ' User's password as SHA1 Hash
        Private m_bIsSignedIn As Boolean = False         ' A flag that determines whether or not the user is signed in.


        '-----------------------------------------------------------------------------------------------------------
        ' C O N S T R U C T O R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' UserInfo Constructor
        ' This constructor is required for de-/serialization which will be done when saving and loading the 
        ' application's settings.
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New()
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' UserInfo Constructor
        ' Creates a new UserInfo instance representing the specified user with the specified password.
        '
        ' Parameters:
        '      userName:   The name of the user.
        '      password:   The user's password.
        '-----------------------------------------------------------------------------------------------------------
        Public Sub New(ByVal userName As String, ByVal password As String)
            m_strName = userName
            m_rbPasswordHash = ComputeSHA1Hash(password)
            m_bIsSignedIn = True
        End Sub

        '-----------------------------------------------------------------------------------------------------------
        ' P R O P E R T I E S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Name Property
        ' Gets or sets the name of the user that is represented by this UserInfo instance.
        '-----------------------------------------------------------------------------------------------------------
        Public Property Name As String
            Get
                Return m_strName
            End Get
            Set(ByVal value As String)
                m_strName = value.Trim()
            End Set
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' PasswordHash Property
        ' Gets or sets the hashed password.
        '-----------------------------------------------------------------------------------------------------------
        Public Property PasswordHash As Byte()
            Get
                Return m_rbPasswordHash
            End Get
            Set(ByVal value As Byte())
                m_rbPasswordHash = value
            End Set
        End Property

        '-----------------------------------------------------------------------------------------------------------
        ' IsSignedIn Property
        ' Gets or sets a value indicating whether the user is signed in or not.
        '-----------------------------------------------------------------------------------------------------------
        Friend Property IsSignedIn As Boolean
            Get
                Return m_bIsSignedIn
            End Get
            Set(ByVal value As Boolean)
                m_bIsSignedIn = value
            End Set
        End Property


        '-----------------------------------------------------------------------------------------------------------
        ' O V E R R I D D E N   M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' Equals Method (overridden)
        ' Equals the object. This instance and the passed object are equal if the object is a UserInfo and 
        ' the names are case insensitive equal.
        '-----------------------------------------------------------------------------------------------------------
        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            If obj Is Nothing Then Return False
            Dim that = TryCast(obj, UserInfo)
            Return Equals(that)
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' GetHashCode Method (overridden)
        '-----------------------------------------------------------------------------------------------------------
        Public Overrides Function GetHashCode() As Integer
            Return MyBase.GetHashCode()
        End Function


        '-----------------------------------------------------------------------------------------------------------
        ' M E T H O D S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' ValidatePassword Method
        ' Validates whether the passed password is equal to the set password of this user info by comparing
        ' the SHA1 hashs.
        '
        ' Parameters:
        '      password:   The password to validate.
        '
        ' Returns: True, if the specified password is correct. Otherwise false.
        '-----------------------------------------------------------------------------------------------------------
        Public Function ValidatePassword(ByVal password As String) As Boolean
            Return Not String.IsNullOrEmpty(password) AndAlso m_rbPasswordHash.SequenceEqual(ComputeSHA1Hash(password))
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' ComputeSHA1Hash Method
        ' Computes a SHA1 hash from the specified text.
        '
        ' Parameters:
        '      text:   The text that is used to compute a SHA1 hash.
        '
        ' Returns: The computed SHA1 hash from the specified text.
        '-----------------------------------------------------------------------------------------------------------
        Private Function ComputeSHA1Hash(ByVal text As String) As Byte()
            If String.IsNullOrEmpty(text) Then Return New Byte(-1) {}
            Dim sha1 As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider()
            Return sha1.ComputeHash(Encoding.UTF8.GetBytes(text))
        End Function

        '-----------------------------------------------------------------------------------------------------------
        ' Equals Method
        ' This instance and the passed UserInfo are equal if the names are case insensitive equal.
        '
        ' Parameters:
        '      user:   The UserInfo instance to compare with.
        ' 
        '  Returns: True if this instance equals to the specified UserInfo object. Otherwise false.
        '-----------------------------------------------------------------------------------------------------------
        Public Overloads Function Equals(ByVal user As UserInfo) As Boolean
            Return user IsNot Nothing AndAlso Name.Equals(user.Name, StringComparison.OrdinalIgnoreCase)
        End Function


        '-----------------------------------------------------------------------------------------------------------
        ' O P E R A T O R S
        '-----------------------------------------------------------------------------------------------------------

        '-----------------------------------------------------------------------------------------------------------
        ' ==  operator
        ' These instances are equal if the names are case insensitive equal.
        '-----------------------------------------------------------------------------------------------------------
        Public Shared Operator =(ByVal a As UserInfo, ByVal b As UserInfo) As Boolean
            Return ReferenceEquals(a, b) OrElse a.Equals(b)
        End Operator

        '-----------------------------------------------------------------------------------------------------------
        ' !=  operator
        ' These instances are not equal if the names are not case insensitive equal.
        '-----------------------------------------------------------------------------------------------------------
        Public Shared Operator <>(ByVal a As UserInfo, ByVal b As UserInfo) As Boolean
            Return Not a Is b
        End Operator
    End Class
End Namespace
