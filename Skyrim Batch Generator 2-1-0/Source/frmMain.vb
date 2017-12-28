'David Kurtz
'Batch File Generator for Skyrim
'Last Revised: 3/27/16
'Current Released Version: 2.1.0
'Current BETA Version: 2.1.0
'Working on: 

Public Class frmMain
    Private Sub BatchNameInputValidation(sender As Object, e As KeyPressEventArgs) Handles txtBatchName.KeyPress
        Select Case e.KeyChar
            Case "A" To "Z"
                'Good input
            Case "a" To "z"
                'Good input
            Case "0" To "9"
                'Good input
            Case vbBack
                'Good input
            Case "_"
                'Good input
            Case "-"
                'Good input
            Case Else
                e.Handled = True
        End Select
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        Dim Answer As Integer = MsgBox("Close the program?", MsgBoxStyle.YesNo)

        If Answer = MsgBoxResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub mnuFileStartOver_Click(sender As Object, e As EventArgs) Handles mnuFileStartOver.Click
        Dim Answer As Integer = MsgBox("Erase preview contents and start over?", MsgBoxStyle.YesNo)

        If Answer = MsgBoxResult.Yes Then
            txtPreview.Text = String.Empty
            ListView.Items.Clear()
        End If
    End Sub

    Private Sub mnuFileCreateBatch_Click(sender As Object, e As EventArgs) Handles mnuFileCreateBatch.Click
        Dim strName As String
        Dim inFile As IO.StreamWriter

        'Check to make sure user entered a name for the batch file
        If txtBatchName.Text = String.Empty Then
            MsgBox("You need to name the Batch File something!")
        Else
            strName = txtBatchName.Text & ".txt"

            'Check to see if a file already exists with the same name
            If IO.File.Exists(strName) = True Then
                'File already exists; prompt to overwrite

                Dim intResult As Integer
                intResult = MsgBox("A file named " & strName & " already exists. Do you want to overwrite this file?", MsgBoxStyle.YesNo)

                If intResult = MsgBoxResult.Yes Then
                    'User wants to overwrite the file; Create Batch File
                    inFile = IO.File.CreateText(strName)

                    'Write Preview Contents to Batch File
                    inFile.Write(txtPreview.Text)

                    'Close Batch File
                    inFile.Close()

                    'Check to make sure batch file was successfully created
                    If IO.File.Exists(strName) Then
                        MsgBox(strName & " successfully created!")
                    Else
                        MsgBox(strName & " was NOT created successfully. Check to make sure you have admin privileges and try again. If the error persists, file a bug report on the Nexus.")
                    End If
                End If
            Else
                'File doesn't already exist; Confirm if they want to Create the Batch File
                Dim Answer As Integer = MsgBox("Create Batch File named '" & strName & "'?", MsgBoxStyle.YesNo)

                If Answer = MsgBoxResult.Yes Then
                    'Create Batch File
                    inFile = IO.File.CreateText(strName)

                    'Write Preview Contents to Batch File
                    inFile.Write(txtPreview.Text)

                    'Close Batch File
                    inFile.Close()

                    'Check to make sure batch file was successfully created
                    If IO.File.Exists(strName) Then
                        MsgBox(strName & " successfully created!")
                    Else
                        MsgBox(strName & " was NOT created successfully. Check to make sure you have admin privileges and try again. If the error persists, file a bug report on the Nexus.")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub mnuModifiersCharacterlevel_Click(sender As Object, e As EventArgs) Handles mnuModifiersCharlevel.Click
        Dim strCharacterLevel As String = InputBox("Please enter your desired character level (must be a whole number greater than zero):", "Character Level")
        Dim intCharacterLevel As Integer

        Integer.TryParse(strCharacterLevel, intCharacterLevel)

        If intCharacterLevel < 1 Then
            'Bad and/or no input, do nothing
        Else
            'Check to see if this is the first entry in txtPreview
            If txtPreview.Text = String.Empty Then
                'It is, so don't include a new line break
                txtPreview.Text = "player.setlevel " & intCharacterLevel
            Else
                'It isn't, so DO include the new line break
                txtPreview.Text &= vbNewLine & "player.setlevel " & intCharacterLevel
            End If

            'Write info to the User's Log
            Dim lvwCurrent As ListViewItem
            lvwCurrent = ListView.Items.Add("player.setlevel " & intCharacterLevel)
            lvwCurrent.SubItems.Add("Sets your character level to " & intCharacterLevel)
        End If
    End Sub

    'Health/Magicka/Stamina click events - utilizes 'sender.text'

    Private Sub HealthMagStamClickEvents(sender As Object, e As EventArgs) Handles mnuModifiersHealth.Click, mnuModifiersMagicka.Click, mnuModifiersStamina.Click
        Dim strValue As String = InputBox("Please enter your desired " & sender.text & " total (must be a whole number greater than zero):", sender.text)
        Dim intValue As Integer
        Dim strInternalName As String = "ERROR: strInternalName HAS NULL VALUE" 'user should never see this

        Integer.TryParse(strValue, intValue)

        If intValue < 1 Then
            'bad input, do nothing
        Else
            'choose an appropriate Internal name for the attribute to be changed
            Select Case sender.text
                Case mnuModifiersHealth.Text
                    strInternalName = "health"
                Case mnuModifiersMagicka.Text
                    strInternalName = "magicka"
                Case mnuModifiersStamina.Text
                    strInternalName = "stamina"
            End Select

            'Check to see if this is the first entry in txtPreview
            If txtPreview.Text = String.Empty Then
                'DON'T include line break
                txtPreview.Text = "player.forceav " & strInternalName & " " & intValue
            Else
                'INCLUDE line break
                txtPreview.Text &= vbNewLine & "player.forceav " & strInternalName & " " & intValue
            End If

            'Write info to the User's Log
            Dim lvwCurrent As ListViewItem
            lvwCurrent = ListView.Items.Add("player.forceav " & strInternalName & " " & intValue)
            lvwCurrent.SubItems.Add("Sets your " & strInternalName & " to " & intValue)
        End If
    End Sub

    'Gold add/remove click events

    Private Sub AddGoldToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuModifiersGoldAdd.Click
        Dim strGold As String = InputBox("Please enter the desired amount of gold to be added (must be a whole number greater than zero):", "Add Gold")
        Dim intGold As Integer

        Integer.TryParse(strGold, intGold)

        If intGold < 1 Then
            'Bad and/or no input, do nothing
        Else
            'Check to see if this is the first entry in txtPreview
            If txtPreview.Text = String.Empty Then
                'It is, so don't include a new line break
                txtPreview.Text = "player.additem f " & intGold
            Else
                'It isn't, so DO include the new line break
                txtPreview.Text &= vbNewLine & "player.additem f " & intGold
            End If

            'Write info to the User's Log
            Dim lvwCurrent As ListViewItem
            lvwCurrent = ListView.Items.Add("player.additem f " & intGold)
            lvwCurrent.SubItems.Add("Adds " & intGold & " gold to your inventory")
        End If
    End Sub

    Private Sub mnuModifiersGoldRemove_Click(sender As Object, e As EventArgs) Handles mnuModifiersGoldRemove.Click
        Dim strGold As String = InputBox("Please enter the desired amount of gold to be removed (must be a whole number greater than zero):", "Remove Gold")
        Dim intGold As Integer

        Integer.TryParse(strGold, intGold)

        If intGold < 1 Then
            'Bad and/or no input, do nothing
        Else
            'Check to see if this is the first entry in txtPreview
            If txtPreview.Text = String.Empty Then
                'It is, so don't include a new line break
                txtPreview.Text = "player.removeitem f " & intGold
            Else
                'It isn't, so DO include the new line break
                txtPreview.Text &= vbNewLine & "player.removeitem f " & intGold
            End If

            'Write info to the User's Log
            Dim lvwCurrent As ListViewItem
            lvwCurrent = ListView.Items.Add("player.removeitem f " & intGold)
            lvwCurrent.SubItems.Add("Removes " & intGold & " gold from your inventory")
        End If
    End Sub

    'Skills click events

    Private Sub SkillsClickEvents(sender As Object, e As EventArgs) Handles mnuModifiersSkillsSpecificCombatArchery.Click, mnuModifiersSkillsGroupsAll.Click,
        mnuModifiersSkillsGroupsCombat.Click, mnuModifiersSkillsGroupsMagic.Click, mnuModifiersSkillsGroupsStealth.Click, mnuModifiersSkillsSpecificCombatBlock.Click,
        mnuModifiersSkillsSpecificCombatHeavyarmor.Click, mnuModifiersSkillsSpecificCombatOnehanded.Click, mnuModifiersSkillsSpecificCombatSmithing.Click, mnuModifiersSkillsSpecificCombatTwohanded.Click,
        mnuModifiersSkillsSpecificMagicAlteration.Click, mnuModifiersSkillsSpecificMagicConjuration.Click, mnuModifiersSkillsSpecificMagicDestruction.Click, mnuModifiersSkillsSpecificMagicEnchanting.Click,
        mnuModifiersSkillsSpecificMagicIllusion.Click, mnuModifiersSkillsSpecificMagicRestoration.Click, mnuModifiersSkillsSpecificStealthAlchemy.Click, mnuModifiersSkillsSpecificStealthLightarmor.Click,
        mnuModifiersSkillsSpecificStealthLockpicking.Click, mnuModifiersSkillsSpecificStealthPickpocket.Click, mnuModifiersSkillsSpecificStealthSneak.Click, mnuModifiersSkillsSpecificStealthSpeech.Click

        'First row = Combat, Second row = Stealth, Third row = Magic
        Dim strSkillNameInternalArray() As String = {"marksman", "block", "heavyarmor", "onehanded", "smithing", "twohanded",
                                                     "alchemy", "lightarmor", "lockpicking", "pickpocket", "sneak", "speechcraft",
                                                     "alteration", "conjuration", "destruction", "enchanting", "illusion", "restoration"}
        Dim strSkillNameInternal As String = "ERROR: strSkillNameInternal HAS NULL VALUE" 'User should never see this
        Dim strSkillValue As String = InputBox("Please enter the value you want " & sender.text & " set to (must be a whole number greater than zero):", sender.text)
        Dim intSkillValue, intArrayStartValue, intArrayEndValue As Integer

        Integer.TryParse(strSkillValue, intSkillValue)

        'Check for valid input
        If intSkillValue < 1 Then
            'bad input, do nothing
        Else
            'Get the appropriate Internal skill name based on which External name tag was passed, OR
            'Set the Loop Boundary variables to the appropriate values
            Select Case sender.text
                Case mnuModifiersSkillsGroupsAll.Text
                    intArrayStartValue = 0
                    intArrayEndValue = 17
                Case mnuModifiersSkillsGroupsCombat.Text
                    intArrayStartValue = 0
                    intArrayEndValue = 5
                Case mnuModifiersSkillsGroupsStealth.Text
                    intArrayStartValue = 6
                    intArrayEndValue = 11
                Case mnuModifiersSkillsGroupsMagic.Text
                    intArrayStartValue = 12
                    intArrayEndValue = 17
                Case mnuModifiersSkillsSpecificCombatArchery.Text
                    strSkillNameInternal = strSkillNameInternalArray(0)
                Case mnuModifiersSkillsSpecificCombatBlock.Text
                    strSkillNameInternal = strSkillNameInternalArray(1)
                Case mnuModifiersSkillsSpecificCombatHeavyarmor.Text
                    strSkillNameInternal = strSkillNameInternalArray(2)
                Case mnuModifiersSkillsSpecificCombatOnehanded.Text
                    strSkillNameInternal = strSkillNameInternalArray(3)
                Case mnuModifiersSkillsSpecificCombatSmithing.Text
                    strSkillNameInternal = strSkillNameInternalArray(4)
                Case mnuModifiersSkillsSpecificCombatTwohanded.Text
                    strSkillNameInternal = strSkillNameInternalArray(5)
                Case mnuModifiersSkillsSpecificStealthAlchemy.Text
                    strSkillNameInternal = strSkillNameInternalArray(6)
                Case mnuModifiersSkillsSpecificStealthLightarmor.Text
                    strSkillNameInternal = strSkillNameInternalArray(7)
                Case mnuModifiersSkillsSpecificStealthLockpicking.Text
                    strSkillNameInternal = strSkillNameInternalArray(8)
                Case mnuModifiersSkillsSpecificStealthPickpocket.Text
                    strSkillNameInternal = strSkillNameInternalArray(9)
                Case mnuModifiersSkillsSpecificStealthSneak.Text
                    strSkillNameInternal = strSkillNameInternalArray(10)
                Case mnuModifiersSkillsSpecificStealthSpeech.Text
                    strSkillNameInternal = strSkillNameInternalArray(11)
                Case mnuModifiersSkillsSpecificMagicAlteration.Text
                    strSkillNameInternal = strSkillNameInternalArray(12)
                Case mnuModifiersSkillsSpecificMagicConjuration.Text
                    strSkillNameInternal = strSkillNameInternalArray(13)
                Case mnuModifiersSkillsSpecificMagicDestruction.Text
                    strSkillNameInternal = strSkillNameInternalArray(14)
                Case mnuModifiersSkillsSpecificMagicEnchanting.Text
                    strSkillNameInternal = strSkillNameInternalArray(15)
                Case mnuModifiersSkillsSpecificMagicIllusion.Text
                    strSkillNameInternal = strSkillNameInternalArray(16)
                Case mnuModifiersSkillsSpecificMagicRestoration.Text
                    strSkillNameInternal = strSkillNameInternalArray(17)
            End Select

            'Set appropriate skill(s) to the user-defined value
            Select Case intArrayEndValue
                Case 0
                    'User only wants one specific skill; check to see if this is first entry in txtPreview
                    If txtPreview.Text = String.Empty Then
                        'DO NOT include line break
                        txtPreview.Text = "player.setav " & strSkillNameInternal & " " & intSkillValue
                    Else
                        'INCLUDE line break
                        txtPreview.Text &= vbNewLine & "player.setav " & strSkillNameInternal & " " & intSkillValue
                    End If

                    'Write info to the User's Log
                    Dim lvwCurrent As ListViewItem
                    lvwCurrent = ListView.Items.Add("player.setav " & strSkillNameInternal & " " & intSkillValue)
                    lvwCurrent.SubItems.Add("Sets your " & sender.text & " skill to level " & intSkillValue)
                Case Else
                    'User wants a group or all skills; check to see if this is first entry in txtPreview
                    If txtPreview.Text = String.Empty Then
                        'DO NOT include line break
                        txtPreview.Text = "player.setav " & strSkillNameInternalArray(intArrayStartValue) & " " & intSkillValue
                    Else
                        'INCLUDE line break
                        txtPreview.Text &= vbNewLine & "player.setav " & strSkillNameInternalArray(intArrayStartValue) & " " & intSkillValue
                    End If

                    'Write info to the User's Log
                    Dim lvwCurrent As ListViewItem
                    lvwCurrent = ListView.Items.Add("player.setav " & strSkillNameInternalArray(intArrayStartValue) & " " & intSkillValue)
                    lvwCurrent.SubItems.Add("One of the lines of code that sets " & sender.text & " to level " & intSkillValue)

                    intArrayStartValue += 1

                    'loop through the array to fill out the appropriate selections
                    For intIncrementor As Integer = intArrayStartValue To intArrayEndValue
                        txtPreview.Text &= vbNewLine & "player.setav " & strSkillNameInternalArray(intIncrementor) & " " & intSkillValue

                        'Write info to the User's Log
                        lvwCurrent = ListView.Items.Add("player.setav " & strSkillNameInternalArray(intIncrementor) & " " & intSkillValue)
                        lvwCurrent.SubItems.Add("One of the lines of code that sets " & sender.text & " to level " & intSkillValue)
                    Next
            End Select
        End If
    End Sub

    'Perks click events

    Private Sub PerksClickEvents(sender As Object, e As EventArgs) Handles mnuModifiersPerksCombatArcheryOverdrawRank1.Click, mnuModifiersPerksCombatArcheryOverdrawRank2.Click,
        mnuModifiersPerksCombatArcheryBullseye.Click, mnuModifiersPerksCombatArcheryCriticalshotRank1.Click, mnuModifiersPerksCombatArcheryCriticalshotRank2.Click,
        mnuModifiersPerksCombatArcheryCriticalshotRank3.Click, mnuModifiersPerksCombatArcheryEagleeye.Click, mnuModifiersPerksCombatArcheryHuntersdiscipline.Click,
        mnuModifiersPerksCombatArcheryOverdrawRank3.Click, mnuModifiersPerksCombatArcheryOverdrawRank4.Click, mnuModifiersPerksCombatArcheryOverdrawRank5.Click,
        mnuModifiersPerksCombatArcheryPowershot.Click, mnuModifiersPerksCombatArcheryQuickshot.Click, mnuModifiersPerksCombatArcheryRanger.Click,
        mnuModifiersPerksCombatArcherySteadyhandRank1.Click, mnuModifiersPerksCombatArcherySteadyhandRank2.Click, mnuModifiersPerksCombatBlockingBlockrunner.Click,
        mnuModifiersPerksCombatBlockingDeadlybash.Click, mnuModifiersPerksCombatBlockingDeflectarrows.Click, mnuModifiersPerksCombatBlockingDisarmingbash.Click,
        mnuModifiersPerksCombatBlockingElementalprotection.Click, mnuModifiersPerksCombatBlockingPowerbash.Click, mnuModifiersPerksCombatBlockingQuickreflexes.Click,
        mnuModifiersPerksCombatBlockingShieldcharge.Click, mnuModifiersPerksCombatBlockingShieldwallRank1.Click, mnuModifiersPerksCombatBlockingShieldwallRank2.Click,
        mnuModifiersPerksCombatBlockingShieldwallRank3.Click, mnuModifiersPerksCombatBlockingShieldwallRank4.Click, mnuModifiersPerksCombatBlockingShieldwallRank5.Click,
        mnuModifiersPerksCombatHeavyarmorConditioning.Click, mnuModifiersPerksCombatHeavyarmorCushioned.Click, mnuModifiersPerksCombatHeavyarmorFistsofsteel.Click,
        mnuModifiersPerksCombatHeavyarmorJuggernautRank1.Click, mnuModifiersPerksCombatHeavyarmorJuggernautRank2.Click, mnuModifiersPerksCombatHeavyarmorJuggernautRank3.Click,
        mnuModifiersPerksCombatHeavyarmorJuggernautRank4.Click, mnuModifiersPerksCombatHeavyarmorJuggernautRank5.Click, mnuModifiersPerksCombatHeavyarmorMatchingset.Click,
        mnuModifiersPerksCombatHeavyarmorReflectblows.Click, mnuModifiersPerksCombatHeavyarmorTowerofstrength.Click, mnuModifiersPerksCombatHeavyarmorWellfitted.Click,
        mnuModifiersPerksCombatOnehandedArmsmanRank1.Click, mnuModifiersPerksCombatOnehandedArmsmanRank2.Click, mnuModifiersPerksCombatOnehandedArmsmanRank3.Click,
        mnuModifiersPerksCombatOnehandedArmsmanRank4.Click, mnuModifiersPerksCombatOnehandedArmsmanRank5.Click, mnuModifiersPerksCombatOnehandedBladesmanRank1.Click,
        mnuModifiersPerksCombatOnehandedBladesmanRank2.Click, mnuModifiersPerksCombatOnehandedBladesmanRank3.Click, mnuModifiersPerksCombatOnehandedBonebreakerRank1.Click,
        mnuModifiersPerksCombatOnehandedBonebreakerRank2.Click, mnuModifiersPerksCombatOnehandedBonebreakerRank3.Click, mnuModifiersPerksCombatOnehandedCriticalcharge.Click,
        mnuModifiersPerksCombatOnehandedDualflurryRank1.Click, mnuModifiersPerksCombatOnehandedDualflurryRank2.Click, mnuModifiersPerksCombatOnehandedDualsavagery.Click,
        mnuModifiersPerksCombatOnehandedFightingstance.Click, mnuModifiersPerksCombatOnehandedHackandslashRank1.Click, mnuModifiersPerksCombatOnehandedHackandslashRank2.Click,
        mnuModifiersPerksCombatOnehandedHackandslashRank3.Click, mnuModifiersPerksCombatOnehandedParalyzingstrike.Click, mnuModifiersPerksCombatOnehandedSavagestrike.Click,
        mnuModifiersPerksCombatSmithingAdvanced.Click, mnuModifiersPerksCombatSmithingArcane.Click, mnuModifiersPerksCombatSmithingDaedric.Click, mnuModifiersPerksCombatSmithingDragon.Click,
        mnuModifiersPerksCombatSmithingDwarven.Click, mnuModifiersPerksCombatSmithingEbony.Click, mnuModifiersPerksCombatSmithingElven.Click, mnuModifiersPerksCombatSmithingGlass.Click,
        mnuModifiersPerksCombatSmithingOrcish.Click, mnuModifiersPerksCombatSmithingSteel.Click, mnuModifiersPerksCombatTwohandedBarbarianRank1.Click,
        mnuModifiersPerksCombatTwohandedBarbarianRank2.Click, mnuModifiersPerksCombatTwohandedBarbarianRank3.Click, mnuModifiersPerksCombatTwohandedBarbarianRank4.Click,
        mnuModifiersPerksCombatTwohandedBarbarianRank5.Click, mnuModifiersPerksCombatTwohandedChampionstance.Click, mnuModifiersPerksCombatTwohandedDeepwoundsRank1.Click,
        mnuModifiersPerksCombatTwohandedDeepwoundsRank2.Click, mnuModifiersPerksCombatTwohandedDeepwoundsRank3.Click, mnuModifiersPerksCombatTwohandedDevastatingblow.Click,
        mnuModifiersPerksCombatTwohandedGreatcritcharge.Click, mnuModifiersPerksCombatTwohandedLimbsplitterRank1.Click, mnuModifiersPerksCombatTwohandedLimbsplitterRank2.Click,
        mnuModifiersPerksCombatTwohandedLimbsplitterRank3.Click, mnuModifiersPerksCombatTwohandedSkullcrusherRank1.Click, mnuModifiersPerksCombatTwohandedSkullcrusherRank2.Click,
        mnuModifiersPerksCombatTwohandedSkullcrusherRank3.Click, mnuModifiersPerksCombatTwohandedSweep.Click, mnuModifiersPerksCombatTwohandedWarmaster.Click,
        mnuModifiersPerksStealthAlchemyAlchemistRank1.Click, mnuModifiersPerksStealthAlchemyAlchemistRank2.Click, mnuModifiersPerksStealthAlchemyAlchemistRank3.Click,
        mnuModifiersPerksStealthAlchemyAlchemistRank4.Click, mnuModifiersPerksStealthAlchemyAlchemistRank5.Click, mnuModifiersPerksStealthAlchemyBenefactor.Click,
        mnuModifiersPerksStealthAlchemyConcpoison.Click, mnuModifiersPerksStealthAlchemyExperimenterRank1.Click, mnuModifiersPerksStealthAlchemyExperimenterRank2.Click,
        mnuModifiersPerksStealthAlchemyExperimenterRank3.Click, mnuModifiersPerksStealthAlchemyGreenthumb.Click, mnuModifiersPerksStealthAlchemyPhysician.Click,
        mnuModifiersPerksStealthAlchemyPoisoner.Click, mnuModifiersPerksStealthAlchemyPurity.Click, mnuModifiersPerksStealthAlchemySnakeblood.Click,
        mnuModifiersPerksStealthLightarmorAgiledefenderRank1.Click, mnuModifiersPerksStealthLightarmorAgiledefenderRank2.Click, mnuModifiersPerksStealthLightarmorAgiledefenderRank3.Click,
        mnuModifiersPerksStealthLightarmorAgiledefenderRank4.Click, mnuModifiersPerksStealthLightarmorAgiledefenderRank5.Click, mnuModifiersPerksStealthLightarmorCustomfit.Click,
        mnuModifiersPerksStealthLightarmorDeftmove.Click, mnuModifiersPerksStealthLightarmorMatchset.Click, mnuModifiersPerksStealthLightarmorUnhindered.Click, mnuModifiersPerksStealthLightarmorWind.Click,
        mnuModifiersPerksStealthLockpickingAdept.Click, mnuModifiersPerksStealthLockpickingApprentice.Click, mnuModifiersPerksStealthLockpickingExpert.Click, mnuModifiersPerksStealthLockpickingGoldtouch.Click,
        mnuModifiersPerksStealthLockpickingLocksmith.Click, mnuModifiersPerksStealthLockpickingMaster.Click, mnuModifiersPerksStealthLockpickingNovice.Click, mnuModifiersPerksStealthLockpickingQuickhands.Click,
        mnuModifiersPerksStealthLockpickingTreasure.Click, mnuModifiersPerksStealthLockpickingUnbreakable.Click, mnuModifiersPerksStealthLockpickingWaxkey.Click, mnuModifiersPerksStealthPickpocketCutpurse.Click,
        mnuModifiersPerksStealthPickpocketExtrapockets.Click, mnuModifiersPerksStealthPickpocketKeymaster.Click, mnuModifiersPerksStealthPickpocketLightfingersRank1.Click, mnuModifiersPerksStealthPickpocketLightfingersRank2.Click,
        mnuModifiersPerksStealthPickpocketLightfingersRank3.Click, mnuModifiersPerksStealthPickpocketLightfingersRank4.Click, mnuModifiersPerksStealthPickpocketLightfingersRank5.Click, mnuModifiersPerksStealthPickpocketMisdirection.Click,
        mnuModifiersPerksStealthPickpocketNight.Click, mnuModifiersPerksStealthPickpocketPerfecttouch.Click, mnuModifiersPerksStealthPickpocketPoisoned.Click, mnuModifiersPerksStealthSneakAssblade.Click,
        mnuModifiersPerksStealthSneakBackstab.Click, mnuModifiersPerksStealthSneakDeadlyaim.Click, mnuModifiersPerksStealthSneakLightfoot.Click, mnuModifiersPerksStealthSneakMufflemove.Click, mnuModifiersPerksStealthSneakShadow.Click,
        mnuModifiersPerksStealthSneakSilence.Click, mnuModifiersPerksStealthSneakSilentroll.Click, mnuModifiersPerksStealthSneakStealthRank1.Click, mnuModifiersPerksStealthSneakStealthRank2.Click, mnuModifiersPerksStealthSneakStealthRank3.Click,
        mnuModifiersPerksStealthSneakStealthRank4.Click, mnuModifiersPerksStealthSneakStealthRank5.Click, mnuModifiersPerksStealthSpeechAllure.Click, mnuModifiersPerksStealthSpeechBribe.Click, mnuModifiersPerksStealthSpeechFence.Click,
        mnuModifiersPerksStealthSpeechHaggleRank1.Click, mnuModifiersPerksStealthSpeechHaggleRank2.Click, mnuModifiersPerksStealthSpeechHaggleRank3.Click, mnuModifiersPerksStealthSpeechHaggleRank4.Click, mnuModifiersPerksStealthSpeechHaggleRank5.Click,
        mnuModifiersPerksStealthSpeechIntimidation.Click, mnuModifiersPerksStealthSpeechInvestor.Click, mnuModifiersPerksStealthSpeechMaster.Click, mnuModifiersPerksStealthSpeechMerchant.Click, mnuModifiersPerksStealthSpeechPersuasion.Click,
        mnuModifiersPerksMagicAlterationAdept.Click, mnuModifiersPerksMagicAlterationApprentice.Click, mnuModifiersPerksMagicAlterationAtronach.Click, mnuModifiersPerksMagicAlterationDual.Click, mnuModifiersPerksMagicAlterationExpert.Click,
        mnuModifiersPerksMagicAlterationMagearmorRank1.Click, mnuModifiersPerksMagicAlterationMagearmorRank2.Click, mnuModifiersPerksMagicAlterationMagearmorRank3.Click, mnuModifiersPerksMagicAlterationMagicresistRank1.Click,
        mnuModifiersPerksMagicAlterationMagicresistRank2.Click, mnuModifiersPerksMagicAlterationMagicresistRank3.Click, mnuModifiersPerksMagicAlterationMaster.Click, mnuModifiersPerksMagicAlterationNovice.Click,
        mnuModifiersPerksMagicAlterationStability.Click, mnuModifiersPerksMagicConjurationAdept.Click, mnuModifiersPerksMagicConjurationApprentice.Click, mnuModifiersPerksMagicConjurationAtro.Click,
        mnuModifiersPerksMagicConjurationDark.Click, mnuModifiersPerksMagicConjurationDual.Click, mnuModifiersPerksMagicConjurationExpert.Click, mnuModifiersPerksMagicConjurationMaster.Click, mnuModifiersPerksMagicConjurationMysticbind.Click,
        mnuModifiersPerksMagicConjurationNecro.Click, mnuModifiersPerksMagicConjurationNovice.Click, mnuModifiersPerksMagicConjurationOblivbind.Click, mnuModifiersPerksMagicConjurationPotency.Click, mnuModifiersPerksMagicConjurationSoulsteal.Click,
        mnuModifiersPerksMagicConjurationSummonerRank1.Click, mnuModifiersPerksMagicConjurationSummonerRank2.Click, mnuModifiersPerksMagicConjurationTwin.Click, mnuModifiersPerksMagicDestructionAdept.Click, mnuModifiersPerksMagicDestructionApprentice.Click,
        mnuModifiersPerksMagicDestructionAugmentflameRank1.Click, mnuModifiersPerksMagicDestructionAugmentflameRank2.Click, mnuModifiersPerksMagicDestructionAugmentfrostRank1.Click, mnuModifiersPerksMagicDestructionAugmentfrostRank2.Click,
        mnuModifiersPerksMagicDestructionAugmentshockRank1.Click, mnuModifiersPerksMagicDestructionAugmentshockRank2.Click, mnuModifiersPerksMagicDestructionDeepfreeze.Click, mnuModifiersPerksMagicDestructionDisintigrate.Click,
        mnuModifiersPerksMagicDestructionDualcast.Click, mnuModifiersPerksMagicDestructionExpert.Click, mnuModifiersPerksMagicDestructionImpact.Click, mnuModifiersPerksMagicDestructionIntenseflame.Click, mnuModifiersPerksMagicDestructionMaster.Click,
        mnuModifiersPerksMagicDestructionNovice.Click, mnuModifiersPerksMagicDestructionRune.Click, mnuModifiersPerksMagicEnchantingCorpus.Click, mnuModifiersPerksMagicEnchantingEnchanterRank1.Click, mnuModifiersPerksMagicEnchantingEnchanterRank2.Click,
        mnuModifiersPerksMagicEnchantingEnchanterRank3.Click, mnuModifiersPerksMagicEnchantingEnchanterRank4.Click, mnuModifiersPerksMagicEnchantingEnchanterRank5.Click, mnuModifiersPerksMagicEnchantingExtra.Click, mnuModifiersPerksMagicEnchantingFire.Click,
        mnuModifiersPerksMagicEnchantingFrost.Click, mnuModifiersPerksMagicEnchantingInsight.Click, mnuModifiersPerksMagicEnchantingSoulsiphon.Click, mnuModifiersPerksMagicEnchantingSoulsqueeze.Click, mnuModifiersPerksMagicEnchantingStorm.Click,
        mnuModifiersPerksMagicIllusionAdept.Click, mnuModifiersPerksMagicIllusionAnimage.Click, mnuModifiersPerksMagicIllusionApprentice.Click, mnuModifiersPerksMagicIllusionAspectterror.Click, mnuModifiersPerksMagicIllusionDual.Click,
        mnuModifiersPerksMagicIllusionExpert.Click, mnuModifiersPerksMagicIllusionHypnotic.Click, mnuModifiersPerksMagicIllusionKindred.Click, mnuModifiersPerksMagicIllusionMaster.Click, mnuModifiersPerksMagicIllusionMastermind.Click,
        mnuModifiersPerksMagicIllusionNovice.Click, mnuModifiersPerksMagicIllusionQuiet.Click, mnuModifiersPerksMagicIllusionRage.Click, mnuModifiersPerksMagicRestorationAdept.Click, mnuModifiersPerksMagicRestorationApprentice.Click,
        mnuModifiersPerksMagicRestorationAvoiddeath.Click, mnuModifiersPerksMagicRestorationDual.Click, mnuModifiersPerksMagicRestorationExpert.Click, mnuModifiersPerksMagicRestorationMaster.Click, mnuModifiersPerksMagicRestorationNecro.Click,
        mnuModifiersPerksMagicRestorationNovice.Click, mnuModifiersPerksMagicRestorationRecoveryRank1.Click, mnuModifiersPerksMagicRestorationRecoveryRank2.Click, mnuModifiersPerksMagicRestorationRegen.Click, mnuModifiersPerksMagicRestorationRespite.Click,
        mnuModifiersPerksMagicRestorationWardabsorb.Click

        'Check to see if this is the first entry in txtPreview
        If txtPreview.Text = String.Empty Then
            'DON'T include line break
            txtPreview.Text = "player.addperk " & sender.tag
        Else
            'INCLUDE line break
            txtPreview.Text &= vbNewLine & "player.addperk " & sender.tag
        End If

        'Write info to the User's Log
        Dim lvwCurrent As ListViewItem
        lvwCurrent = ListView.Items.Add("player.addperk " & sender.tag)
        lvwCurrent.SubItems.Add("Grants you the '" & sender.text & "' Perk")
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        AboutBox.Show()
    End Sub

    Private Sub InstructionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles mnuHelpInstructions.Click
        Dim frmInstructions As frmInstructions = frmInstructions.Instance

        frmInstructions.Show()
    End Sub

    'Stone Spells Click events

    Private Sub StonesClickEvents(sender As Object, e As EventArgs) Handles mnuModifiersStoneApprentice.Click, mnuModifiersStoneAtronach.Click, mnuModifiersStoneLady.Click, mnuModifiersStoneLord.Click, mnuModifiersStoneLover.Click,
        mnuModifiersStoneMage.Click, mnuModifiersStoneRitual.Click, mnuModifiersStoneSerpent.Click, mnuModifiersStoneShadow.Click, mnuModifiersStoneSteed.Click, mnuModifiersStoneThief.Click, mnuModifiersStoneTower.Click, mnuModifiersStoneWarrior.Click

        'Check to see if this is the first entry in txtPreview
        If txtPreview.Text = String.Empty Then
            'DON'T include line break
            txtPreview.Text = "player.addspell " & sender.tag
        Else
            'INCLUDE line break
            txtPreview.Text &= vbNewLine & "player.addspell " & sender.tag
        End If

        'Write info to the User's Log
        Dim lvwCurrent As ListViewItem
        lvwCurrent = ListView.Items.Add("player.addspell " & sender.tag)
        lvwCurrent.SubItems.Add("Gives you the " & sender.text & " Stone power")
    End Sub

    'Modes Click events

    Private Sub mnuModesStandard_Click(sender As Object, e As EventArgs) Handles mnuModesStandard.Click
        'Check to see if Standard mode is currently active
        If mnuModesStandard.Checked = True Then
            'Currently active; do nothing
        Else
            'Warn user that switching modes erases preview contents
            Dim Answer As Integer = MsgBox("Switching modes erases Preview contents. Switch to Standard Mode?", MsgBoxStyle.YesNo)

            If Answer = MsgBoxResult.Yes Then
                'Set Standard mode to active and Advanced to inactive
                mnuModesStandard.Checked = True
                mnuModesAdvanced.Checked = False

                'Show Standard view to user and hide advanced view
                ListView.Visible = True
                txtPreview.Visible = False

                'Clear both modes of content
                ListView.Items.Clear()
                txtPreview.Text = String.Empty
            End If
        End If
    End Sub

    Private Sub mnuModesAdvanced_Click(sender As Object, e As EventArgs) Handles mnuModesAdvanced.Click
        'Check to see if Advanced mode is currently active
        If mnuModesAdvanced.Checked = True Then
            'Currently active; do nothing
        Else
            'Warn user that switching modes erases preview contents
            Dim Answer As Integer = MsgBox("Switching modes erases Preview contents. Switch to Advanced Mode?", MsgBoxStyle.YesNo)

            If Answer = MsgBoxResult.Yes Then
                'Set Advanced mode to active and Standard to inactive
                mnuModesStandard.Checked = False
                mnuModesAdvanced.Checked = True

                'Show Advanced view to user and hide Standard view
                ListView.Visible = False
                txtPreview.Visible = True

                'Clear both modes of content
                ListView.Items.Clear()
                txtPreview.Text = String.Empty
            End If
        End If
    End Sub

    'Shouts click events

    Private Sub mnuModifiersShoutsAnimal_Click(sender As Object, e As EventArgs) Handles mnuModifiersShoutsAnimal.Click, mnuModifiersShoutsAura.Click, mnuModifiersShoutsCalldragon.Click,
        mnuModifiersShoutsCallvalor.Click, mnuModifiersShoutsClearskies.Click, mnuModifiersShoutsDisarm.Click, mnuModifiersShoutsDismay.Click, mnuModifiersShoutsDragonrend.Click,
        mnuModifiersShoutsElementalfury.Click, mnuModifiersShoutsEthereal.Click, mnuModifiersShoutsFirebreath.Click, mnuModifiersShoutsForce.Click, mnuModifiersShoutsFrostbreath.Click,
        mnuModifiersShoutsIceform.Click, mnuModifiersShoutsKynes.Click, mnuModifiersShoutsMarked.Click, mnuModifiersShoutsSlow.Click, mnuModifiersShoutsSprint.Click,
        mnuModifiersShoutsStorm.Click, mnuModifiersShoutsThrow.Click

        Dim strWord1 As String = "ERROR: strWord1 HAS NULL VALUE"
        Dim strWord2 As String = "ERROR: strWord2 HAS NULL VALUE"
        Dim strWord3 As String = "ERROR: strWord3 HAS NULL VALUE"

        'Get the appropriate shout codes
        Select Case sender.text
            Case mnuModifiersShoutsAnimal.Text
                strWord1 = "60291"
                strWord2 = "60292"
                strWord3 = "60293"
            Case mnuModifiersShoutsAura.Text
                strWord1 = "60294"
                strWord2 = "60295"
                strWord3 = "60296"
            Case mnuModifiersShoutsEthereal.Text
                strWord1 = "32917"
                strWord2 = "32918"
                strWord3 = "32919"
            Case mnuModifiersShoutsCalldragon.Text
                strWord1 = "46b89"
                strWord2 = "46b8a"
                strWord3 = "46b8b"
            Case mnuModifiersShoutsCallvalor.Text
                strWord1 = "51960"
                strWord2 = "51961"
                strWord3 = "51962"
            Case mnuModifiersShoutsClearskies.Text
                strWord1 = "3cd31"
                strWord2 = "3cd32"
                strWord3 = "3cd33"
            Case mnuModifiersShoutsDisarm.Text
                strWord1 = "5fb95"
                strWord2 = "5fb96"
                strWord3 = "5fb97"
            Case mnuModifiersShoutsDismay.Text
                strWord1 = "3291a"
                strWord2 = "3291b"
                strWord3 = "3291c"
            Case mnuModifiersShoutsDragonrend.Text
                strWord1 = "44251"
                strWord2 = "44252"
                strWord3 = "44253"
            Case mnuModifiersShoutsElementalfury.Text
                strWord1 = "3291d"
                strWord2 = "3291e"
                strWord3 = "3291f"
            Case mnuModifiersShoutsFirebreath.Text
                strWord1 = "20e17"
                strWord2 = "20e18"
                strWord3 = "20e19"
            Case mnuModifiersShoutsFrostbreath.Text
                strWord1 = "5d16c"
                strWord2 = "5d16d"
                strWord3 = "5d16e"
            Case mnuModifiersShoutsIceform.Text
                strWord1 = "602a3"
                strWord2 = "602a4"
                strWord3 = "602a5"
            Case mnuModifiersShoutsKynes.Text
                strWord1 = "6029d"
                strWord2 = "6029e"
                strWord3 = "6029f"
            Case mnuModifiersShoutsMarked.Text
                strWord1 = "60297"
                strWord2 = "60298"
                strWord3 = "60299"
            Case mnuModifiersShoutsSlow.Text
                strWord1 = "48aca"
                strWord2 = "48acb"
                strWord3 = "48acc"
            Case mnuModifiersShoutsStorm.Text
                strWord1 = "6029a"
                strWord2 = "6029b"
                strWord3 = "6029c"
            Case mnuModifiersShoutsThrow.Text
                strWord1 = "602a0"
                strWord2 = "602a1"
                strWord3 = "602a2"
            Case mnuModifiersShoutsForce.Text
                strWord1 = "13e22"
                strWord2 = "13e23"
                strWord3 = "13e24"
            Case mnuModifiersShoutsSprint.Text
                strWord1 = "2f7bb"
                strWord2 = "2f7bc"
                strWord3 = "2f7bd"
        End Select

        'Check to see if this is the first entry in txtpreview
        If txtPreview.Text = String.Empty Then
            'DON'T include line break
            txtPreview.Text = "player.teachword " & strWord1 & vbNewLine & "player.unlockword " & strWord1 & vbNewLine &
                "player.teachword " & strWord2 & vbNewLine & "player.unlockword " & strWord2 & vbNewLine & "player.teachword " &
                strWord3 & vbNewLine & "player.unlockword " & strWord3
        Else
            'DO include line break
            txtPreview.Text &= vbNewLine & "player.teachword " & strWord1 & vbNewLine & "player.unlockword " & strWord1 & vbNewLine &
                "player.teachword " & strWord2 & vbNewLine & "player.unlockword " & strWord2 & vbNewLine & "player.teachword " &
                strWord3 & vbNewLine & "player.unlockword " & strWord3
        End If

        'Write info to the User's log
        Dim lvwCurrent As ListViewItem
        'Word 1
        lvwCurrent = ListView.Items.Add("player.teachword " & strWord1)
        lvwCurrent.SubItems.Add("Teaches you Word 1 of the " & sender.text & " Shout")
        lvwCurrent = ListView.Items.Add("player.unlockword " & strWord1)
        lvwCurrent.SubItems.Add("Unlocks Word 1 of the " & sender.text & " Shout")
        'Word 2
        lvwCurrent = ListView.Items.Add("player.teachword " & strWord2)
        lvwCurrent.SubItems.Add("Teaches you Word 2 of the " & sender.text & " Shout")
        lvwCurrent = ListView.Items.Add("player.unlockword " & strWord2)
        lvwCurrent.SubItems.Add("Unlocks Word 2 of the " & sender.text & " Shout")
        'Word 3
        lvwCurrent = ListView.Items.Add("player.teachword " & strWord3)
        lvwCurrent.SubItems.Add("Teaches you Word 3 of the " & sender.text & " Shout")
        lvwCurrent = ListView.Items.Add("player.unlockword " & strWord3)
        lvwCurrent.SubItems.Add("Unlocks Word 3 of the " & sender.text & " Shout")
    End Sub
End Class
