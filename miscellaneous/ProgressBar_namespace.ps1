<#
esempio di come si può usare il namespace (in questo caso per il form di una progess bar)
#>

using namespace System.Windows.Forms
using namespace System.Drawing

Add-Type -AssemblyName System.Windows.Forms

$maxProgressSteps = 10

# Create the form.
$form = [Form] @{
  Text = "TRANSFER RATE"; Size = [Size]::new(600, 200); StartPosition = 'CenterScreen'; TopMost = $true; MinimizeBox = $false; MaximizeBox = $false; FormBorderStyle = 'FixedSingle'
}
# Add controls.
$form.Controls.AddRange(@(
  ($label = [Label] @{ Location = [Point]::new(20, 20); Size = [Size]::new(550, 30) })
  ($bar = [ProgressBar] @{ Location = [Point]::new(20, 70); Size = [Size]::new(550, 30); Style = 'Continuous'; Maximum = $maxProgressSteps })
))

# Start the long-running background job that
# emits objects as they become available.
$job = Start-Job {
  foreach ($i in 1..$using:maxProgressSteps) {
    $i
    Start-Sleep -Milliseconds 500
  }
}

# Show the form *non-modally*, i.e. execution
# of the script continues, and the form is only
# responsive if [System.Windows.Forms.Application]::DoEvents() is called periodically.
$null = $form.Show()

while ($job.State -notin 'Completed', 'Failed') {
  # Check for new output objects from the background job.
  if ($output = Receive-Job $job) {
    $step = $output[-1] # Use the last object output.
    # Update the progress bar.
    $label.Text = '{0} / {1}' -f $step, $maxProgressSteps
    $bar.Value = $step
  }  

  # Allow the form to process events.
  [System.Windows.Forms.Application]::DoEvents()

  # Sleep a little, to avoid a near-tight loop.
  Start-Sleep -Milliseconds 200
}

# Clean up.
$form.Dispose()
Remove-Job $job -Force