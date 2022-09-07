$subList=Get-Content -Path "C:\Users\ko16i50\Downloads\eao-azure\AzureLandscapeInventory\SubscriptionList.txt"

$count=0
$fileName="AzureAllResourceGroup"+(Get-Date -f MM-dd-yyyy_HH_mm_ss)+".csv"
$fileNameWithPath="C:\Users\ko16i50\OneDrive - Kohler Co\Desktop\repos\$fileName"
New-Item $fileNameWithPath -ItemType File
Add-Content -Path $fileNameWithPath -Value "ResourceGroupName,Location,OwnerName,OwnerDept,CostCenter,Environment,Purpose,Subscription"
foreach ($sub in $subList) {
    $subscriptionInfo=Select-AzSubscription -Subscription $sub
    $subscriptionInfo
    $rgList=Get-AzResourceGroup
    foreach ($rg in $rgList) 
    {
    
                if($null -ne $rg.Tags)
                {
                    if($null -ne $rg.Tags["owner"])
                    {
                    $owner=$rg.Tags["owner"]
                    }
                    elseif ($null -ne $rg.Tags["Owner"]) {
                    $owner=$rg.Tags["Owner"] 
                    }
                    else {
                        $owner=$null  
                    }
                    $costCenter=$rg.Tags["costCenter"]
                    $ownerDept=$rg.Tags["ownerDept"]
                    
                    if($null -ne $rg.Tags["environment"])
                    {
                    $environment=$rg.Tags["environment"]
                    }
                    elseif ($null -ne $rg.Tags["Environment"]) {
                    $environment=$rg.Tags["Environment"] 
                    }
                    else {
                        $environment=$null  
                    }
                    
                    if($null -ne $rg.Tags["purpose"])
                    {
                    $purpose=$rg.Tags["purpose"]
                    }
                    elseif ($null -ne $rg.Tags["Purpose"]) {
                    $purpose=$rg.Tags["Purpose"] 
                    }
                    else {
                        $purpose=$null  
                    }
                }
                else {
                    $owner=$null
                    $ownerDept=$null
                    $costCenter=$null
                    $environment=$null
                    $purpose=$null  
                }
                $content=($rg.ResourceGroupName+","+$rg.Location+","+$owner+","+$ownerDept+","+$costCenter+","+$environment+","+$purpose+","+$subscriptionInfo.Subscription.Name)
                Add-Content -Path $fileNameWithPath -Value $content
                $count++
         
    } 
    Write-Host "Number of ResourceGroup "$rgList.Count
}