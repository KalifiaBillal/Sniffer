Option Strict On
Imports System.IO
Imports System.Text
Imports System.ComponentModel
Imports System.EventArgs
Imports System.Net.NetworkInformation
Imports System.Net


Public Class Form6
    Private Sub Form7_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim proc As New System.Diagnostics.Process()
        Dim path_ As String = Application.StartupPath & "\Result.txt"
        proc.StartInfo.RedirectStandardOutput = True ' on redirige le flux de la console
        proc.StartInfo.UseShellExecute = False ' pour que le flux arrive vers notre textbox
        proc.StartInfo.CreateNoWindow = True ' true pour ne pas avoir la fenetre noire, false pour l'avoir
        proc.StartInfo.FileName = "cmd.exe"
        proc.StartInfo.Arguments = "/c netsh wlan show networks mode=bssid"
        proc.Start()
        RichTextBox1.Text = proc.StandardOutput.ReadToEnd ' on affiche le flux que nous avons préalablement redirigé
        RichTextBox1.Text = RichTextBox1.Text.Replace("ÿ", " ")
        RichTextBox1.Text = RichTextBox1.Text.Replace("r‚seaux", "réseaux")
        RichTextBox1.Text = RichTextBox1.Text.Replace("r‚seau", "réseau")
        System.IO.File.WriteAllText(path_, RichTextBox1.Text) 'ce code va remettre le fichier à zéro à chaque lancement du programme
        System.IO.File.AppendAllText(path_, RichTextBox1.Text) ' ce code va enregistrer les nouveaux résultats à la suite des précédents
    End Sub
    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '
        Call RefreshList()
        '
    End Sub
    '
    Private Sub RefreshList()
        '
        cbo_Interfaces.Items.Clear()
        '
        For Each Device As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces
            cbo_Interfaces.Items.Add(Device.Name)
        Next
        '
        cbo_Interfaces.Enabled = cbo_Interfaces.Items.Count <> 0
        cbo_Interfaces.SelectedIndex = CInt(IIf(cbo_Interfaces.Items.Count > 0, 0, -1))
        '
    End Sub
    '
    Private Sub cbo_Interfaces_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbo_Interfaces.SelectedIndexChanged
        '
        FillNetworkInfos(cbo_Interfaces.SelectedIndex)
        '
    End Sub
    '
    Private Sub FillNetworkInfos(ByVal Index As Integer)
        '
        If Index <> -1 Then
            '
            Dim Device As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces(Index)
            '
            lbl_MacAdress.Text = "Adresse physique: " & Tools.FormatMACAddress(Device.GetPhysicalAddress.ToString)
            lbl_Desc.Text = "Description: " & Device.Description
            lbl_Type.Text = "Type d'interface: " & Device.NetworkInterfaceType.ToString()
            lbl_Status.Text = "Statut: " & Tools.StatusConverter(Device.OperationalStatus)
            lbl_Speed.Text = "Vitesse: " & Tools.SizeConverter(Device.Speed) & "bits/s"
            '
            lv_Stats.Items.Clear()
            With lv_Stats.Items.Add("Octet(s) reçu(s)")
                .SubItems.Add(Tools.SizeConverter(Device.GetIPv4Statistics.BytesReceived) & "o (" & Device.GetIPv4Statistics.BytesReceived.ToString & ")")
                .Group = lv_Stats.Groups("InputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) protocole inconnu")
                .SubItems.Add(Device.GetIPv4Statistics.IncomingUnknownProtocolPackets.ToString)
                .Group = lv_Stats.Groups("InputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) erroné(s)")
                .SubItems.Add(Device.GetIPv4Statistics.IncomingPacketsWithErrors.ToString)
                .Group = lv_Stats.Groups("InputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) non unicast")
                .SubItems.Add(Device.GetIPv4Statistics.NonUnicastPacketsReceived.ToString)
                .Group = lv_Stats.Groups("InputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) unicast")
                .SubItems.Add(Device.GetIPv4Statistics.UnicastPacketsSent.ToString)
                .Group = lv_Stats.Groups("InputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) ignoré(s)")
                .SubItems.Add(Device.GetIPv4Statistics.IncomingPacketsDiscarded.ToString)
                .Group = lv_Stats.Groups("InputTransfert")
            End With
            '
            With lv_Stats.Items.Add("Octet(s) envoyé(s)")
                .SubItems.Add(Tools.SizeConverter(Device.GetIPv4Statistics.BytesSent) & "o (" & Device.GetIPv4Statistics.BytesSent.ToString & ")")
                .Group = lv_Stats.Groups("OutputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) erroné(s)")
                .SubItems.Add(Device.GetIPv4Statistics.OutgoingPacketsWithErrors.ToString)
                .Group = lv_Stats.Groups("OutputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) non unicast")
                .SubItems.Add(Device.GetIPv4Statistics.NonUnicastPacketsSent.ToString)
                .Group = lv_Stats.Groups("OutputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) unicast")
                .SubItems.Add(Device.GetIPv4Statistics.UnicastPacketsSent.ToString)
                .Group = lv_Stats.Groups("OutputTransfert")
            End With
            With lv_Stats.Items.Add("Paquet(s) ignoré(s)")
                .SubItems.Add(Device.GetIPv4Statistics.OutgoingPacketsDiscarded.ToString)
                .Group = lv_Stats.Groups("OutputTransfert")
            End With
            With lv_Stats.Items.Add("Taille de la file d'attente")
                .SubItems.Add(Device.GetIPv4Statistics.OutputQueueLength.ToString)
                .Group = lv_Stats.Groups("OutputTransfert")
            End With
            '
            lst_IPmc.Items.Clear()
            For Each AddressInfos As IPAddressInformation In Device.GetIPProperties.MulticastAddresses
                If AddressInfos.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then lst_IPmc.Items.Add(AddressInfos.Address.ToString)
            Next
            '
            lst_IPuc.Items.Clear()
            For Each AddressInfos As IPAddressInformation In Device.GetIPProperties.UnicastAddresses
                If AddressInfos.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then lst_IPuc.Items.Add(AddressInfos.Address.ToString)
            Next
            '
            lst_DHCP.Items.Clear()
            For Each Address As IPAddress In Device.GetIPProperties.DhcpServerAddresses
                If Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then lst_DHCP.Items.Add(Address.ToString)
            Next
            '
            lst_Gateway.Items.Clear()
            For Each Gateway As GatewayIPAddressInformation In Device.GetIPProperties.GatewayAddresses
                If Gateway.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then lst_Gateway.Items.Add(Gateway.Address.ToString)
            Next
            '
            lst_DNS.Items.Clear()
            For Each Address As IPAddress In Device.GetIPProperties.DnsAddresses
                If Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then lst_DNS.Items.Add(Address.ToString)
            Next
            '
            lst_WINS.Items.Clear()
            For Each Address As IPAddress In Device.GetIPProperties.WinsServersAddresses
                If Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then lst_WINS.Items.Add(Address.ToString)
            Next
            '
        End If
        '
    End Sub
End Class