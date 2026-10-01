. ./mock2.ps1
function Get-MgRoleManagementDirectoryRoleDefinition { param([switch]$All,$Property,$ErrorAction)
  @([pscustomobject]@{Id='r-ga';DisplayName='Global Administrator'},[pscustomobject]@{Id='r-gr';DisplayName='Global Reader'},[pscustomobject]@{Id='r-ex';DisplayName='Exchange Administrator'}) }
function Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance { param([switch]$All,$ErrorAction)
  @([pscustomobject]@{RoleDefinitionId='r-ga';PrincipalId='u3';AssignmentType='Assigned';EndDateTime=$null},
    [pscustomobject]@{RoleDefinitionId='r-ga';PrincipalId='u1';AssignmentType='Activated';EndDateTime=(Get-Date).AddHours(4)}) }
function Get-MgRoleManagementDirectoryRoleAssignment { param([switch]$All,$ErrorAction)
  @([pscustomobject]@{RoleDefinitionId='r-ga';PrincipalId='u3';DirectoryScopeId='/'},[pscustomobject]@{RoleDefinitionId='r-ga';PrincipalId='u1';DirectoryScopeId='/'},
    [pscustomobject]@{RoleDefinitionId='r-gr';PrincipalId='g1aaaaaaaa';DirectoryScopeId='/'},[pscustomobject]@{RoleDefinitionId='r-gr';PrincipalId='sp-1';DirectoryScopeId='/'}) }
function Get-MgRoleManagementDirectoryRoleEligibilitySchedule { param([switch]$All,$ErrorAction)
  @([pscustomobject]@{RoleDefinitionId='r-ga';PrincipalId='u1';DirectoryScopeId='/'},[pscustomobject]@{RoleDefinitionId='r-ex';PrincipalId='u2';DirectoryScopeId='/administrativeUnits/au1'}) }
function Get-MgGroupOwner { param($GroupId,[switch]$All,$ErrorAction)
  if($GroupId -eq 'g3cccccccc'){ @([pscustomobject]@{Id='u4';AdditionalProperties=@{}}) } elseif($GroupId -eq 'g4dddddddd'){ @() } else { @([pscustomobject]@{Id='u1';AdditionalProperties=@{}}) } }
