#   ---
#
#    Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
#    The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
#    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#
#	---


Function Invoke-AI {
<#
.SYNOPSIS
 Invoke Artificial Insanity. As this is not very intelligent, yet.
.DESCRIPTION
 Ask questions, receive answers. Tell the bot if it was correct or not. You need to train it.
 You can add questions and answers as training data to $datapath\data.txt in this format: question;answer;1
 You can answer to an unanswered question by sending message "train" to the bot
 The training data is used to guess answers. More training data leads to more accurate answers.
.EXAMPLE
 You: My phone doesn't work
 Bot: Have you tried turning it off and on again
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [string]$message,
        [string]$datapath,
        [string]$sender
    )

    $tempfolder   = "$datapath\temp"             # Temp is used to temporarily store previous communication
    $specialpath  = "$datapath\special.txt"      # Special words to increase answer accuracy
    $newQuestions = "$datapath\newquestions.txt" # Unanswered questions are kept here
    $AIdata       = "$datapath\data.txt"         # Training data or sort of "knowledge base", bot's answers are based on this

    if (!(Test-Path $AIdata)) # Test if datafile exists and create if it doesn't
    {
        "your question;you need to add questions and answers to data.txt;1" | Out-File -FilePath $AIdata -Encoding utf8
    }

    if (!(Test-Path $specialpath)) # Test if specialfile exists and create if it doesn't
    {
        "" | Out-File -FilePath $specialpath -Encoding utf8
    }

    if (!(Test-Path $tempfolder)) # Test if tempfolder exists and create if it doesn't
    {
        New-Item -ItemType Directory $tempfolder | out-null
    }
    
    $data = import-csv -Delimiter ';' -Encoding UTF8 -Header question,answer,score -Path $AIdata

    $specials = Get-Content $specialpath

    Try 
    {
        $newQuestion = Get-Content $newquestions | select -First 1
    }
    Catch 
    {
        $newQuestion = ""
    }

    if (Test-Path $tempfolder\$sender.tmp) {
        $lastAnswer        = Get-Content $tempfolder\$sender.tmp -Tail 1
        $lastAnswerContent = $lastAnswer.split(';') 
        $question          = $lastAnswerContent | select -Skip 1 -First 1
        $solution          = $lastAnswerContent | select -Skip 2 -First 1
        [int]$lastscore    = $lastAnswerContent | select -Skip 3 -First 1
    }

    if (Test-Path $tempfolder\$sender.tmpq) {
        
        $question = Get-Content $tempfolder\$sender.tmpq

        "$question;$message;5" | Out-File -Append -Encoding utf8 -FilePath $AIdata
        Remove-Item -Path $tempfolder\$sender.tmpq
        $result = "Thanks!"
    
        $newContents = get-content $newquestions | Select-String $question -NotMatch 
        $newContents | Set-Content -Path $newquestions
    }
    elseif ($message -eq "train") {
        if ($newQuestion -ne ""){
            $result = $newQuestion
            $newQuestion | Out-File -Append -Encoding utf8 -FilePath $tempfolder\$sender.tmpq
        }
        else {
            $result = "No more questions"
        }
    }
    elseif ($message -eq "yes" -and (Test-Path $tempfolder\$sender.tmp)) # answer was correct and question & answer pair is saved to knowledge base with positive score
    {
        $newscore = $lastscore

        $newDataContents = get-content $AIdata | Select-String "$question;$solution" -NotMatch 
        $newDataContents | Set-Content -Path $AIdata

        "$question;$solution;$newscore"  | Out-File -Append -Encoding utf8 -FilePath $AIdata
        Remove-Item -Path $tempfolder\$sender.tmp

        $newContents = get-content $newQuestions | Select-String $question -NotMatch 
        $newContents | Set-Content -Path $newquestions
        $result = "Great! :)"
    }
    elseif ($message -eq "no" -and (Test-Path $tempfolder\$sender.tmp)) # answer was incorrect and question & answer pair is saved to knowledge base with 0 score
    {                    
        $newscore = 0
        $question | Out-File -Append -Encoding utf8 -FilePath $newQuestions

        $newDataContents = get-content $AIdata | Select-String "$question;$solution" -NotMatch 
        $newDataContents | Set-Content -Path $AIdata
        
        "$question;$solution;$newscore"| Out-File -Append -Encoding utf8 -FilePath $AIdata
        Remove-Item -Path $tempfolder\$sender.tmp
        $result = "Ok my bad..."
    }
    else 
    {
        $msgWords = $message.Split(' ')
        $answer_list = @()

        foreach ($item in $data)
        {
            $score = 0

            foreach ($word in $msgWords) # go through the words in user message
            {
                if ($item.question.contains("$word")) # if question contains same words as message -> score +1
                {
                    $score++

                    if ($specials -eq "$word") # if question has special words -> score +4 / special
                        {
                            $score += 4
                        }
                }
                
            } # foreach msgWords

        
            $totalScore = ($score * $item.score) + 1 # calculate total score based on knowledgebase score, special words and matching words

            $properties = @{
                'answer'   = $item.answer
                'value'    = $totalScore
            }

            $entry = New-Object -TypeName psobject -Property $properties 

            $newEntry = $true

            if ($newEntry)
            {
                $answer_list += $entry
            }

        } # foreach data

        $answer = $answer_list | sort value -Descending | select -first 1 # sort answers by score and select the one with the highest score
        
        $result = $answer.answer 
        
        $newdata = $sender+";"+$message+";"+$result+";"+$answer.value
        $newdata | Out-File -Append -Encoding utf8 -FilePath $tempfolder\$sender.tmp
        $result = $result+"`n"+"Did this solve your issue? (yes/no)"

    } # else

    return $result
}