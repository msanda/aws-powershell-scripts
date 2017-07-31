# Function to creat the EC2 Instance worksheet
function Create-EC2InstanceWorksheet {

        Write-Host "Creating EC2 Instances Worksheet..`n`n" -ForegroundColor Green

        # Adding worksheet
        $workbook.Worksheets.Add()

        # Creating the worksheet for Virtual Machine
        $VirtualMachineWorksheet = $workbook.Worksheets.Item(1)
        $VirtualMachineWorksheet.Name = 'VirtualMachine'

        # Headers for the worksheet
        $VirtualMachineWorksheet.Cells.Item(1,1) = 'Region'
        $VirtualMachineWorksheet.Cells.Item(1,2) = 'VM Name'
        $VirtualMachineWorksheet.Cells.Item(1,3) = 'VM Image ID'
        $VirtualMachineWorksheet.Cells.Item(1,4) = 'VM Instance ID'
        $VirtualMachineWorksheet.Cells.Item(1,5) = 'VM Instance Type'
        $VirtualMachineWorksheet.Cells.Item(1,6) = 'VM Private IP'
        $VirtualMachineWorksheet.Cells.Item(1,7) = 'VM Public IP'
        $VirtualMachineWorksheet.Cells.Item(1,8) = 'VM VPC ID'
        $VirtualMachineWorksheet.Cells.Item(1,9) = 'VM Subnet ID'
        $VirtualMachineWorksheet.Cells.Item(1,10) = 'VM State'
        $VirtualMachineWorksheet.Cells.Item(1,11) = 'VM Security Group Id'
        
        # Excel Cell Counter
        $row_counter = 3
        $column_counter = 1


    # Get the Ec2 instances for each region
    foreach($AWS_Locations_Iterator in $AWS_Locations){
        $EC2Instances = Get-EC2Instance -Region $AWS_Locations_Iterator

        # Iterating over each instance
        foreach($EC2Instances_Iterator in $EC2Instances){
            
            # Ignore if a region does not have any instances
            if($EC2Instances_Iterator.count -eq $null) {
            continue
            }
            # Populating the cells
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $AWS_Locations_Iterator
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.keyname.tostring()
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.imageid
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.Instanceid
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.Instancetype.Value
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.PrivateIpAddress
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.PublicIpAddress
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.vpcid
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.SubnetId
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.state.name.value
            $VirtualMachineWorksheet.Cells.Item($row_counter,$column_counter++) = $EC2Instances_Iterator.Instances.securitygroups.GroupId

            # Seting the row and column counter for next EC2 instance entry
            $row_counter = $row_counter + 1
            $column_counter = 1
        }

        # Iterating to the next region
        $row_counter = $row_counter + 3
    }

}