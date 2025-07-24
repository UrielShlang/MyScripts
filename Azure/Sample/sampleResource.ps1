$rg = 'arm-introduction-01'
New-AzResourceGroup -Name $rg -Location northeurope -Force

New-AzResourceGroupDeployment `
    -Name 'new-storage' `
    -ResourceGroupName $rg `
    -TemplateFile '01-storage.json' `
    -TemplateParameterFile '.\04-storage-function.parameters.json' 
    ## Parameters
    -StorageName 'amdemostorageintro02'
