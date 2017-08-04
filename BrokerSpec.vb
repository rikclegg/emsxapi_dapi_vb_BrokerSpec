' Copyright 2017. Bloomberg Finance L.P.
'
' Permission Is hereby granted, free of charge, to any person obtaining a copy
' of this software And associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, And/Or
' sell copies of the Software, And to permit persons to whom the Software Is
' furnished to do so, subject to the following conditions:  The above
' copyright notice And this permission notice shall be included in all copies
' Or substantial portions of the Software.
'
' THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS Or
' IMPLIED, INCLUDING BUT Not LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS Or COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES Or OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT Or OTHERWISE, ARISING
' FROM, OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS
' IN THE SOFTWARE.
'

Imports Bloomberglp.Blpapi

Namespace com.bloomberg.emsx.samples
    Module BrokerSpec

        Private ReadOnly SESSION_STARTED As New Name("SessionStarted")
        Private ReadOnly SESSION_STARTUP_FAILURE As New Name("SessionStartupFailure")
        Private ReadOnly SERVICE_OPENED As New Name("ServiceOpened")
        Private ReadOnly SERVICE_OPEN_FAILURE As New Name("ServiceOpenFailure")
        Private ReadOnly ERROR_INFO As New Name("ErrorInfo")
        Private ReadOnly BROKER_SPEC As New Name("BrokerSpec")

        Private d_service As String
        Private d_host As String
        Private d_port As Integer

        Private quit As Boolean = False

        Private requestID As CorrelationID

        Sub Main(ByVal args As String())

            System.Console.WriteLine("Bloomberg - EMSX API Example - BrokerSpec")

            Dim example As BrokerSpec = New BrokerSpec()
            example.run()

            Do
            Loop Until quit

            'System.Console.ReadLine()

        End Sub

        Class BrokerSpec

            Sub New()
                'The BrokerSpec service Is only available in the production environment
                d_service = "//blp/emsx.brokerspec"
                d_host = "localhost"
                d_port = 8194
            End Sub

            Friend Sub run()

                Dim d_sessionOptions As SessionOptions
                Dim session As Session

                d_sessionOptions = New SessionOptions()
                d_sessionOptions.ServerHost = d_host
                d_sessionOptions.ServerPort = d_port

                session = New Session(d_sessionOptions, New EventHandler(AddressOf processEvent))
                session.StartAsync()

            End Sub

            Public Sub processEvent(ByVal eventObj As [Event], ByVal session As Session)

                Try
                    Select Case eventObj.Type

                        Case [Event].EventType.SESSION_STATUS
                            processSessionEvent(eventObj, session)
                            Exit Select
                        Case [Event].EventType.SERVICE_STATUS
                            processServiceEvent(eventObj, session)
                            Exit Select
                        Case [Event].EventType.RESPONSE
                            processResponseEvent(eventObj, session)
                            Exit Select
                        Case Else
                            processMiscEvent(eventObj, session)
                            Exit Select
                    End Select
                Catch ex As Exception
                    System.Console.Error.WriteLine(ex)
                End Try

            End Sub

            Private Sub processSessionEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj
                    If msg.MessageType.Equals(SESSION_STARTED) Then
                        System.Console.WriteLine("Session started...")
                        session.OpenServiceAsync(d_service)
                    ElseIf msg.MessageType.Equals(SESSION_STARTUP_FAILURE) Then
                        System.Console.Error.WriteLine("Error: Session startup failed")
                    End If
                Next msg

            End Sub

            Private Sub processServiceEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj
                    If msg.MessageType.Equals(SERVICE_OPENED) Then
                        System.Console.WriteLine("Service opened...")

                        Dim service As Service = session.GetService(d_service)

                        Dim request As Request = service.CreateRequest("GetBrokerSpecForUuid")

                        request.Set("uuid", 8049857)

                        System.Console.WriteLine("Request: " + request.ToString)

                        requestID = New CorrelationID()

                        ' Submit the request
                        Try
                            session.SendRequest(request, requestID)
                        Catch ex As Exception
                            System.Console.Error.WriteLine("Failed to send the request: " + ex.Message)
                        End Try

                    ElseIf msg.MessageType.Equals(SERVICE_OPEN_FAILURE) Then
                        System.Console.Error.WriteLine("Error: Service failed to open")
                    End If

                Next msg

            End Sub

            Private Sub processResponseEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj

                    System.Console.WriteLine("MESSAGE: " + msg.ToString)
                    System.Console.WriteLine("CORRELATION ID: " + msg.CorrelationID.ToString)

                    If msg.CorrelationID Is requestID Then

                        System.Console.WriteLine("Message Type: " + msg.MessageType.ToString)

                        If msg.MessageType.Equals(ERROR_INFO) Then
                            Dim errorCode As Integer = msg.GetElementAsInt32("ERROR_CODE")
                            Dim errorMessage As String = msg.GetElementAsString("ERROR_MESSAGE")
                            System.Console.WriteLine("ERROR CODE: " + errorCode.ToString + Chr(9) + "ERROR MESSAGE: " + errorMessage)
                        ElseIf msg.MessageType.Equals(BROKER_SPEC) Then
                            Dim brokers As Element = msg.GetElement("brokers")

                            Dim numBkrs As Integer = brokers.NumValues

                            System.Console.WriteLine("Number of Brokers: " + numBkrs.ToString)

                            For i As Integer = 0 To numBkrs - 1

                                Dim broker As Element = brokers.GetValueAsElement(i)

                                Dim code As String = broker.GetElementAsString("code")
                                Dim assetClass As String = broker.GetElementAsString("assetClass")

                                If broker.HasElement("strategyFixTag") Then
                                    Dim tag As Long = broker.GetElementAsInt64("strategyFixTag")
                                    System.Console.WriteLine(Chr(13) + "Broker code: " + code + Chr(9) + "Class: " + assetClass + Chr(9) + "Tag: " + tag.ToString)

                                    Dim strats As Element = broker.GetElement("strategies")

                                    Dim numStrats As Integer = strats.NumValues

                                    System.Console.WriteLine(Chr(9) + "No. of Strategies: " + numStrats.ToString)

                                    For s As Integer = 0 To numStrats - 1

                                        Dim strat As Element = strats.GetValueAsElement(s)

                                        Dim Name As String = strat.GetElementAsString("name")
                                        Dim fixVal As String = strat.GetElementAsString("fixValue")

                                        System.Console.WriteLine(Chr(9) + "Strategy Name: " + Name + Chr(9) + "Fix Value: " + fixVal)

                                        Dim parameters As Element = strat.GetElement("parameters")

                                        Dim numParams As Integer = parameters.NumValues

                                        System.Console.WriteLine(Chr(9) + Chr(9) + "No. of Parameters: " + numParams.ToString)

                                        For p As Integer = 0 To numParams - 1

                                            Dim param As Element = parameters.GetValueAsElement(p)

                                            Dim pname As String = param.GetElementAsString("name")
                                            Dim fixTag As Long = param.GetElementAsInt64("fixTag")
                                            Dim required As Boolean = param.GetElementAsBool("isRequired")
                                            Dim replaceable As Boolean = param.GetElementAsBool("isReplaceable")

                                            System.Console.WriteLine(Chr(9) + Chr(9) + "Parameter: " + pname + Chr(9) + "Tag: " + fixTag.ToString + Chr(9) + "Required: " + required.ToString + Chr(9) + "Replaceable: " + replaceable.ToString)

                                            Dim typeName As String = param.GetElement("type").GetElement(0).Name.ToString()

                                            Dim vals As String = ""

                                            If typeName.Equals("enumeration") Then
                                                Dim enumerators As Element = param.GetElement("type").GetElement(0).GetElement("enumerators")

                                                Dim numEnums As Integer = enumerators.NumValues

                                                For e As Integer = 0 To numEnums - 1
                                                    Dim en As Element = enumerators.GetValueAsElement(e)

                                                    vals = vals + en.GetElementAsString("name") + "[" + en.GetElementAsString("fixValue") + "],"
                                                Next e
                                                vals = vals.Substring(0, vals.Length - 1)
                                            ElseIf typeName.Equals("range") Then
                                                Dim rng As Element = param.GetElement("type").GetElement(0)
                                                Dim mn As Long = rng.GetElementAsInt64("min")
                                                Dim mx As Long = rng.GetElementAsInt64("max")
                                                Dim st As Long = rng.GetElementAsInt64("step")
                                                vals = "min:" + mn.ToString + " max:" + mx.ToString + " step:" + st.ToString
                                            ElseIf typeName.Equals("string") Then
                                                Dim possVals As Element = param.GetElement("type").GetElement(0).GetElement("possibleValues")

                                                Dim numVals As Integer = possVals.NumValues

                                                For v As Integer = 0 To numVals - 1
                                                    vals = vals + possVals.GetValueAsString(v) + ","
                                                Next v
                                                If vals.Length > 0 Then vals = vals.Substring(0, vals.Length - 1)
                                            End If

                                            If vals.Length > 0 Then
                                                System.Console.WriteLine(Chr(9) + Chr(9) + Chr(9) + "Type: " + typeName + " (" + vals + ")")
                                            Else
                                                System.Console.WriteLine(Chr(9) + Chr(9) + Chr(9) + "Type: " + typeName)
                                            End If

                                        Next p
                                    Next s
                                Else
                                    System.Console.WriteLine(Chr(13) + "Broker code: " + code + Chr(9) + "class: " + assetClass)
                                    System.Console.WriteLine(Chr(9) + "No Strategies")
                                End If


                                System.Console.WriteLine(Chr(13) + Chr(9) + "Time In Force: ")
                                Dim tifs As Element = broker.GetElement("timesInForce")
                                Dim numTifs As Integer = tifs.NumValues
                                For t As Integer = 0 To numTifs - 1
                                    Dim tif As Element = tifs.GetValueAsElement(t)
                                    Dim tifName As String = tif.GetElementAsString("name")
                                    Dim tifFixValue As String = tif.GetElementAsString("fixValue")
                                    System.Console.WriteLine(Chr(9) + chr(9) + "Name " + tifName + Chr(9) + "Fix Value: " + tifFixValue)
                                Next t

                                System.Console.WriteLine(Chr(13) + Chr(9) + "Order Types:")

                                Dim ordTypes As Element = broker.GetElement("orderTypes")
                                Dim numOrdTypes As Integer = ordTypes.NumValues

                                For o As Integer = 0 To numOrdTypes - 1
                                    Dim ordType As Element = ordTypes.GetValueAsElement(o)
                                    Dim typName As String = ordType.GetElementAsString("name")
                                    Dim typFixValue As String = ordType.GetElementAsString("fixValue")
                                    System.Console.WriteLine(Chr(9) + Chr(9) + "Name: " + typName + Chr(9) + "Fix Value: " + typFixValue)
                                Next o

                                System.Console.WriteLine(Chr(13) + Chr(9) + "Handling Instructions:")

                                Dim handInsts As Element = broker.GetElement("handlingInstructions")
                                Dim numHandInsts As Integer = handInsts.NumValues

                                For h As Integer = 0 To numHandInsts - 1
                                    Dim handInst As Element = handInsts.GetValueAsElement(h)
                                    Dim instName As String = handInst.GetElementAsString("name")
                                    Dim instFixValue As String = handInst.GetElementAsString("fixValue")
                                    System.Console.WriteLine(Chr(9) + Chr(9) + "Name: " + instName + Chr(9) + "Fix Value: " + instFixValue)
                                Next h
                            Next i
                        End If

                        quit = True
                        session.Stop()
                    End If
                Next msg
            End Sub

            Private Sub processMiscEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj
                    System.Console.WriteLine("MESSAGE: " + msg.ToString)
                Next msg

            End Sub

        End Class

    End Module

End Namespace
