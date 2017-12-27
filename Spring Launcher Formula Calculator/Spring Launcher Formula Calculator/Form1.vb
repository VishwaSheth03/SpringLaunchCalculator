Public Class Form1

    Dim Hp, Ht, d, θ, i, m, k, v As Double
    Const g As Double = 9.8

    Private Sub ProjectileHeightButton_Click(sender As System.Object, e As System.EventArgs) Handles ProjectileHeightButton.Click
        ProjectileHeightLabel.Text = "= " + Str(ProjectileHeightTextBox.Text) + " m"
        Hp = Val(ProjectileHeightTextBox.Text)
    End Sub

    Private Sub TargetHeightButton_Click(sender As System.Object, e As System.EventArgs) Handles TargetHeightButton.Click
        TargetHeightLabel.Text = "= " + Str(TargetHeightTextBox.Text) + " m"
        Ht = Val(TargetHeightTextBox.Text)
    End Sub

    Private Sub HorizontalDistanceButton_Click(sender As System.Object, e As System.EventArgs) Handles HorizontalDistanceButton.Click
        HorizontalDistanceLabel.Text = "= " + Str(HorizontalDistanceTextBox.Text) + " m"
        d = Val(HorizontalDistanceTextBox.Text)
    End Sub

    Private Sub AngleButton_Click(sender As System.Object, e As System.EventArgs) Handles AngleButton.Click
        AngleLabel.Text = "= " + Str(AngleTextBox.Text) + " degrees"
        θ = Val(AngleTextBox.Text)
    End Sub

    Private Sub EquilibriumLenghtButton_Click(sender As System.Object, e As System.EventArgs) Handles EquilibriumLenghtButton.Click
        EquilibriumLengthLabel.Text = "= " + Str(EquilibriumLengthTextBox.Text) + " m"
        i = Val(EquilibriumLengthTextBox.Text)
    End Sub

    Private Sub SpringMassButton_Click(sender As System.Object, e As System.EventArgs) Handles SpringMassButton.Click
        SpringMassLabel.Text = "= " + Str(SpringMassTextBox.Text) + " kg"
        m = Val(SpringMassTextBox.Text)
    End Sub

    Private Sub SpringConstantButton_Click(sender As System.Object, e As System.EventArgs) Handles SpringConstantButton.Click
        SpringConstantLabel.Text = "= " + Str(SpringConstantTextBox.Text) + " N/m"
        k = Val(SpringConstantTextBox.Text)
    End Sub

    Private Sub CalculateButton_Click(sender As System.Object, e As System.EventArgs) Handles CalculateButton.Click
        LaunchSpeedLabel.Text = LaunchSpeed(Hp, Ht, d, θ, i, m, k)
        v = LaunchSpeed(Hp, Ht, d, θ, i, m, k)

        If xPullDistancePlus(Hp, Ht, d, θ, i, m, k, v) > xPullDistanceMinus(Hp, Ht, d, θ, i, m, k, v) Then
            xPullDistanceLabel.Text = xPullDistancePlus(Hp, Ht, d, θ, i, m, k, v) * 100
        Else : xPullDistanceLabel.Text = xPullDistanceMinus(Hp, Ht, d, θ, i, m, k, v) * 100
        End If

    End Sub

    Function LaunchSpeed(ByVal Hp As Double, Ht As Double, d As Double, θ As Double, i As Double, m As Double, k As Double) As Double
        Return Math.Sqrt((g * d ^ 2) / (2 * ((Math.Cos(θ * Math.PI / 180)) ^ 2) * (d * Math.Tan(θ * Math.PI / 180) + Hp - Ht)))
    End Function

    Function xPullDistancePlus(ByVal Hp As Double, Ht As Double, d As Double, θ As Double, i As Double, m As Double, k As Double, v As Double) As Double
        Return ((2 * m * g * Math.Sin(θ * Math.PI / 180)) + Math.Sqrt((4 * (m ^ 2) * (g ^ 2) * (Math.Sin(θ * Math.PI / 180) ^ 2)) + (4 * k * (m * (v ^ 2) + (2 * m * g * Math.Sin(θ * Math.PI / 180) * i))))) / (2 * k)
    End Function

    Function xPullDistanceMinus(ByVal Hp As Double, Ht As Double, d As Double, θ As Double, i As Double, m As Double, k As Double, v As Double) As Double
        Return ((2 * m * g * Math.Sin(θ * Math.PI / 180)) - Math.Sqrt((4 * (m ^ 2) * (g ^ 2) * (Math.Sin(θ * Math.PI / 180) ^ 2)) + (4 * k * (m * (v ^ 2) + (2 * m * g * Math.Sin(θ * Math.PI / 180) * i))))) / (2 * k)
    End Function

End Class
