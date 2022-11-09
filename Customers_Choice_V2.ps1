##############################################################################################
# Script Name  : Customer Choice
# Description  : Provisioning Customer Choice Script
# Developed By : Workplace Innovation Team (Marco Caimi - Gianmario Casula)
# Company      : Elmec Informatica S.p.A.
# Version      : 2.0 With Stack Panel
##############################################################################################

####################################################
#region Nascondi Powershell Console
####################################################
    # .Net methods for hiding/showing the console in the background
    Add-Type -Name Window -Namespace Console -MemberDefinition '
    [DllImport("Kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
    '

    function Show-Console
                                                                {
    $consolePtr = [Console.Window]::GetConsoleWindow()

    # Hide = 0,
    # ShowNormal = 1,
    # ShowMinimized = 2,
    # ShowMaximized = 3,
    # Maximize = 3,
    # ShowNormalNoActivate = 4,
    # Show = 5,
    # Minimize = 6,
    # ShowMinNoActivate = 7,
    # ShowNoActivate = 8,
    # Restore = 9,
    # ShowDefault = 10,
    # ForceMinimized = 11

    [Console.Window]::ShowWindow($consolePtr, 4)
    }

    function Hide-Console
                {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    #0 hide
    [Console.Window]::ShowWindow($consolePtr, 2)
    }
####################################################
#endregion Nascondi Powershell Console
####################################################

####################################################
#region XAML GUI preparation and configuration
####################################################

    $global:ScriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition


    
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    
    [string]$XamlCodeString = @'
        <Window 
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            xmlns:ed="http://schemas.microsoft.com/expression/2010/drawing"
        
                Title="Elmec Informatica" WindowStartupLocation="CenterScreen" Height="400" Width="500" FontFamily="Arial" FontSize="14" WindowStyle="None" AllowsTransparency="True" Background="Green">
    
            <Grid Margin="0,0,0,0" HorizontalAlignment="Right" Width="500">
                <Grid.Background>
                    <ImageBrush ImageSource="#####IMAGESOURCE#####" Stretch="UniformToFill"/>
                </Grid.Background>
                <StackPanel Name="XAML_StackPanel_Step1" Visibility="Visible">
                    <Label Content="Please select a Customer:" FontFamily="Arial Black" FontSize="16" Background="Transparent" Foreground="White" HorizontalAlignment="Center" HorizontalContentAlignment="Center"  VerticalAlignment="Top" Height="30" Width="300" Margin="0,110,0,0"/>
                    <ComboBox Name="XAML_ComboBox1" FontFamily="Arial" FontSize="14" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Margin="0,20,0,0" HorizontalAlignment="Center" VerticalAlignment="Center" Width="300" Height="30"/>
                    <ComboBox Name="XAML_ComboBox2" FontFamily="Arial" FontSize="14" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Margin="0,20,0,0" HorizontalAlignment="Center" VerticalAlignment="Center" Width="300" Height="30"/>
                    <Button Name="XAML_NextPageButton" Content="Next" FontFamily="Arial" FontSize="14" Background="#ffffff" Foreground="Black" BorderThickness="0" FontWeight="Bold" VerticalAlignment="Top" HorizontalAlignment="Center" Height="30" Width="120" Margin="0,20,0,0"/>
                    <Button Name="XAML_CloseButton1" Content="Let's Close!" FontFamily="Arial" FontSize="14" Background="#ffffff" Foreground="Black" BorderThickness="0" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Top" Width="120" Height="30" Margin="350,70,0,0"/>
                </StackPanel>
                <StackPanel Name="XAML_StackPanel_Step2" Visibility="Hidden">
                    <Label Content="Customer Selection Recap:" FontFamily="Arial Black" FontSize="16" Background="Transparent" Foreground="White" HorizontalAlignment="Center" HorizontalContentAlignment="Center" VerticalAlignment="Top" Height="30" Width="300" Margin="0,110,0,0"/>
                    <ListView Name="XAML_BulletPointListView" HorizontalAlignment="Center" VerticalAlignment="Top" Height="160" Width="500" Margin="0,10,0,0">
                        <ListView.Resources>
                            <Style TargetType="GridViewColumnHeader">
                                <Setter Property="Visibility" Value="Collapsed" />
                            </Style>
                            <Style TargetType="ListViewItem">
                                <Setter Property="FontWeight" Value="Bold" />
                                <Setter Property="Foreground" Value="Black" />
                            </Style>
                        </ListView.Resources>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="" Width="235" DisplayMemberBinding="{Binding BulletName}"/><!-- Bullet Point Name -->
                                <GridViewColumn Header="" Width="235" DisplayMemberBinding="{Binding BulletValue}"/><!-- Bullet Point Value -->
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Button Name="XAML_GoButton" Content="Let's Go!" FontFamily="Arial" FontSize="14" Background="#ffffff" Foreground="Black" BorderThickness="0" FontWeight="Bold" VerticalAlignment="Top" HorizontalAlignment="Center" Height="30" Width="120" Margin="0,10,0,0"/>
                    <Button Name="XAML_CloseButton2" Content="Let's Close!" FontFamily="Arial" FontSize="14" Background="#ffffff" Foreground="Black" BorderThickness="0" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Top" Width="120" Height="30" Margin="350,10,0,0"/>
                    <Button Name="XAML_PreviousPageButton" Content="Previous" FontFamily="Arial" FontSize="14" Background="#ffffff" Foreground="Black" BorderThickness="0" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Top" Width="120" Height="30" Margin="-350,-30,0,0"/>
                </StackPanel>
            </Grid>
        </Window>
'@

    #Set Image Source
    $ImageName = "OSDCloud_Elmec.png"
    [xml]$XamlCode = $XamlCodeString.Replace("#####IMAGESOURCE#####","$ScriptPath\$ImageName")

    #Creating GUI Variable from XAML
    $XamlReader = (New-Object System.Xml.XmlNodeReader $XamlCode)
    $GUI = [Windows.Markup.XamlReader]::Load($XamlReader)

    #Creating Variable from XAML
    $XamlCode.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name) }

####################################################
#endregion GUI preparation and configuration
####################################################

####################################################
#region Getting Customer XML Configuration files
####################################################

    #Name of XML File
    #Notes: è importante che sia presente una cartella per ogni cliente, 
    #       dentro ciascuna cartella deve esserci il file OSDConfiguration.xml
    $XMLfile = ‘OSDConfiguration.xml’

    $script:SourceFolder = "$ScriptPath"
    $CustomerFolders = (Get-ChildItem -Path $SourceFolder -Directory).Name

    #Lista che conterrà tutte le info di tutti i clienti
    $CustomersList = New-Object -TypeName 'System.Collections.ArrayList'
        
    foreach ($Customer in $CustomerFolders){
        
        $XMLCustomerPath = "$SourceFolder\$Customer\$XMLfile"

        if (Test-Path $XMLCustomerPath){
            
            #Getting specific Customer XML File
            [XML]$XMLContent = Get-Content $XMLCustomerPath
            
            #Creazione Oggetti CustomerInfo da aggiungere
            $XMLData = new-object -TypeName PSObject
            
            #Nome Cliente
            $XMLData | Add-Member -MemberType NoteProperty -Name CustomerName -Value ($XMLContent.osdconfiguration.CustomerName)

            #Creazione lista per contenere la lista di oggetti "opzione"
            $OptionsList = New-Object -TypeName 'System.Collections.ArrayList'
            $XMLData | Add-Member -MemberType NoteProperty -Name Options -Value $OptionsList

            #Per ogni opzione contenuta nell'xml del cliente
            foreach($XMLContentOption in $XMLContent.OSDConfiguration.Options.Option){

                #Creazione Oggetto Option
                $XMLDataOption = new-object -TypeName PSObject
                
                #Aquisizione Dati della specifica Option per il cliente
                $XMLDataOption | Add-Member -MemberType NoteProperty -Name OptionName -Value ($XMLContentOption.OptionName)
                $XMLDataOption | Add-Member -MemberType NoteProperty -Name WinOS -Value ($XMLContentOption.WinOS)
                $XMLDataOption | Add-Member -MemberType NoteProperty -Name WinBuild -Value ($XMLContentOption.WinBuild)
                $XMLDataOption | Add-Member -MemberType NoteProperty -Name WinEdition -Value ($XMLContentOption.WinEdition)
                $XMLDataOption | Add-Member -MemberType NoteProperty -Name WinLanguage -Value ($XMLContentOption.WinLanguage)
                $XMLDataOption | Add-Member -MemberType NoteProperty -Name ManagementTool -Value ($XMLContentOption.ManagementTool)
                
                #Aggiunta Option nella lista delle opzioni
                $XMLData.Options.add($XMLDataOption)

                }
            
            #Aggiunta Informazioni del cliente nella lista CustomersList
            $CustomersList.Add($XMLData)
            
            }
        }

####################################################
#endregion Getting Customer XML Configuration files
####################################################

####################################################
#region Loading data to GUI
####################################################

    #$XAML_BulletPointListView.Visibility = "Hidden"

    #$XAML_Background.ImageSource ="$ScriptPath\OSDCloud_Elmec_Green.png" 
    #$XAML_Background.ImageSource ="C:\Users\casula\Elmec Informatica S.p.A\WP - OSDCloud - General\OSD Structure\Customers\OSDCloud_Elmec_Green.png" 

    $XAML_NextPageButton.Visibility = "Hidden"
    $XAML_ComboBox2.Visibility = "Hidden"

    foreach($Customer in $CustomersList.CustomerName){
        $XAML_ComboBox1.Items.add($Customer)
        }

####################################################
#endregion Loading data to GUI
####################################################

####################################################
#region Setting GUI Element Events
####################################################
    
    
    #ComboBox Selection changed event
    $XAML_ComboBox1.Add_SelectionChanged({
        
       foreach ($Customer in $CustomersList){

            if ($Customer.CustomerName -eq $XAML_ComboBox1.SelectedItem){
                
                $global:CustomerSelected = $Customer

                
                $XAML_ComboBox2.Text = ""
                $XAML_ComboBox2.Items.Clear()

                foreach($OptionName in $CustomerSelected.Options.OptionName){

                    $XAML_ComboBox2.Items.add($OptionName)
                
                    }
                    

                }
            }

        $XAML_BulletPointListView.Clear()
        
        $XAML_NextPageButton.Visibility = "Hidden"
        $XAML_ComboBox2.Visibility = "Visible"

        $XAML_ComboBox2.SelectedIndex = "0"

        #QUESTI SONO I VALORI CHE SERVONO A OSD CLOUD
        #write-host "$CustomerSelected"
        })

    #ComboBox Selection 2 changed event
    $XAML_ComboBox2.Add_SelectionChanged({
        
        foreach ($CustomerSelectedOption in $CustomerSelected.Options){

            if ($CustomerSelectedOption.OptionName -eq $XAML_ComboBox2.SelectedItem){
                
                #Variabile globale per salvare l'opzione selezionata nella seconda  combobox
                $global:OptionSelected = $CustomerSelected.Options | Where {$_.OptionName -eq $CustomerSelectedOption.OptionName}

                #Creazione Lista Bullet Point da aggiungere alla vista
                $XMLDatainfoList = New-Object -TypeName 'System.Collections.ArrayList'
                
                $XMLData0 = new-object -TypeName PSObject
                $XMLData0 | Add-Member -MemberType NoteProperty -Name BulletName -Value "Customer Name"
                $XMLData0 | Add-Member -MemberType NoteProperty -Name BulletValue -Value ($CustomerSelected.CustomerName)
                $XMLDatainfoList.Add($XMLData0)

                $XMLData1 = new-object -TypeName PSObject
                $XMLData1 | Add-Member -MemberType NoteProperty -Name BulletName -Value "Selected Option"
                $XMLData1 | Add-Member -MemberType NoteProperty -Name BulletValue -Value ($CustomerSelectedOption.OptionName)
                $XMLDatainfoList.Add($XMLData1)

                $XMLData2 = new-object -TypeName PSObject
                $XMLData2 | Add-Member -MemberType NoteProperty -Name BulletName -Value "OS"
                $XMLData2 | Add-Member -MemberType NoteProperty -Name BulletValue -Value ($CustomerSelectedOption.WinOS)
                $XMLDatainfoList.Add($XMLData2)

                $XMLData3 = new-object -TypeName PSObject
                $XMLData3 | Add-Member -MemberType NoteProperty -Name BulletName -Value "Build"
                $XMLData3 | Add-Member -MemberType NoteProperty -Name BulletValue -Value ($CustomerSelectedOption.WinBuild)
                $XMLDatainfoList.Add($XMLData3)

                $XMLData4 = new-object -TypeName PSObject
                $XMLData4 | Add-Member -MemberType NoteProperty -Name BulletName -Value "Edition"
                $XMLData4 | Add-Member -MemberType NoteProperty -Name BulletValue -Value ($CustomerSelectedOption.WinEdition)
                $XMLDatainfoList.Add($XMLData4)

                $XMLData5 = new-object -TypeName PSObject
                $XMLData5 | Add-Member -MemberType NoteProperty -Name BulletName -Value "Language"
                $XMLData5 | Add-Member -MemberType NoteProperty -Name BulletValue -Value ($CustomerSelectedOption.WinLanguage)
                $XMLDatainfoList.Add($XMLData5)

                $XMLData6 = new-object -TypeName PSObject
                $XMLData6 | Add-Member -MemberType NoteProperty -Name BulletName -Value "Management Tool"
                $XMLData6 | Add-Member -MemberType NoteProperty -Name BulletValue -Value ($CustomerSelectedOption.ManagementTool)
                $XMLDatainfoList.Add($XMLData6)
            
                $CustomerSelectedBulletList = $XMLDatainfoList

                }
            }

        $XAML_BulletPointListView.Clear()
        $XAML_BulletPointListView.ItemsSource = $CustomerSelectedBulletList

        #QUESTI SONO I VALORI CHE SERVONO A OSD CLOUD
        write-host "$CustomerSelected"

        $XAML_NextPageButton.Visibility = "Visible"

        })


    #Button Let's Go Click event
    $XAML_GoButton.Add_Click({

        $GUI.WindowState = [System.Windows.Forms.FormWindowState]::Minimized
        Show-Console
        cls
        Write-Host "***********************************************************************" -ForegroundColor Green
        Write-Host "*                                                                     *" -ForegroundColor Green
        Write-Host "*                          Elmec Informatica                          *" -ForegroundColor Green
        Write-Host "*                                                                     *" -ForegroundColor Green
        Write-Host "***********************************************************************" -ForegroundColor Green
        Write-Host " "
        & "$ScriptPath\$($CustomerSelected.CustomerName)\OSDCloud.ps1"
        
        ########################################
        # CALL OSD CLOUD FUNCTION
        ########################################

        })

    #Button Next Page Click event
    $XAML_NextPageButton.Add_Click({

        $XAML_StackPanel_Step1.Visibility = "Hidden"
        $XAML_StackPanel_Step2.Visibility = "Visible"
    })

    #Button Previous Page Click event
    $XAML_PreviousPageButton.Add_Click({

        $XAML_StackPanel_Step1.Visibility = "Visible"
        $XAML_StackPanel_Step2.Visibility = "Hidden"
    })

    #Button Close Click event
    $XAML_CloseButton1.Add_Click({

        $GUI.Close()
    })

    #Button Close Click event
    $XAML_CloseButton2.Add_Click({

        $GUI.Close()
    })

####################################################
#endregion Setting GUI Element Events
####################################################

####################################################
#region Start and Load GUI
###################################################

    $GUI.Topmost = $true
    Hide-Console
    $GUI.ShowDialog() | out-null
    

####################################################
#endregion Start and Load GUI
###################################################