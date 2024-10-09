<#
Name......: Quotes.ps1
Version...: 24.10.1
Author....: Dario CORRADA

Just suggesting remarkable quotes
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# just pipe more than single "Split-Path" if the script maps to nested subfolders
$ThisFile = $myinvocation.MyCommand.Definition

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

$slurp = @()
foreach ($newline in (Get-Content $ThisFile)) {
    if ($newline -match "^<<") {
        $newquote = "$newline"
    } elseif ($newline -match ">>$") {
        $newquote = "$newquote $newline"
        $newquote -match "^<<(.+)>>$" | Out-Null
        $newquote = $matches[1]
        $slurp += $newquote
    }
}

$DnD = (Get-Random -Maximum $slurp.Count) - 1

[System.Windows.MessageBox]::Show("$($slurp[$DnD])",'JOB TERMINATED','Ok','info') | Out-Null

<# THE QUOTES
<<Give them nothing, but take from them everything.
[300]>>
<<Xerxes: It would be nothing short of madness for you, brave king, and your valiant
troops to perish. All because of a simple misunderstanding. There is much our
cultures could share.
Leonidas: Haven't you noticed, we've been sharing our culture with you all morning.
[300]>>
<<Persian: Our arrows will blot out the sun!
Stelios: Then we will fight in the shade.
[300]>>
<<Unless I miss my guess, we're in for one wild night.
[300]>>
<<Spartans! Ready your breakfast and eat hearty... For tonight, we dine in hell!
[300]>>
<<Immortals... they fail our king's test. And a man who fancies himself 
a god feels a very human chill crawl up his spine
[300]>>
<<Remember us. As simple an order as a king can give. For he did not wish tribute, nor
song, or monuments or poems of war and valor. His wish was simple. Remember us, he said
to me. That was his hope, should any free soul come across that place, in all the
countless centuries yet to be. May all our voices whisper to you from the ageless stones
[300]>>
<<Go tell the Spartans, passerby, that here by Spartan law, we lie.
[Simonides' epigram]>>
<<The world will know that free men stood against a tyrant, that few stood against 
many, and that before this battle is over, even a god-king can bleed
[300]>>
<<Hundreds leave. A handful stays. Only one looks back
[300]>>
<<Persion Emissary: This is madness!
Leonidas: Madness? This is SPARTA!
[300]>>
<<Immortals... we put their name to the test
[300]>>
<<You have many slaves, Xerxes, but few warriors. It won't be long before they 
fear my spears more than your whisp
[300]>>
<<His helmet was stifling, it narrowed his vision. And he must see far. His shield was
heavy. It threw him off balance. And his target is far away
[300]>>
<<Prefer punishment to disgraceful gain; for the one is painful but once, but the
other for one's whole life.
[Chilon]>>
<<If one is strong be also merciful, so that one's neighbors may respect one
rather than fear one.
[Chilon]>>
<<Do not let one's tongue outrun one's sense.
[Chilon]>>
<<Do not make too much haste on one's road.
[Chilon]>>
<<Horror has a face... and you must make a friend of horror. Horror and moral terror are 
your friends. If they are not, then they are enemies to be feared. They are truly enemies!
[Apocalypse Now]>>
<<I love the smell of napalm in the morning. [...] The smell, you know that 
gasoline smell, the whole hill. Smelled like victory.
[Apocalypse Now]>>
<<A donkey, a donkey. My kingdom for a donkey!
[Worms Armageddon>>
<<God has a hard-on for Marines because we kill everything we see! He plays His 
games, we play ours! To show our appreciation for so much power, we keep heaven 
packed with fresh souls! God was here before the Marine Corps! So you can give 
your heart to Jesus, but your ass belongs to the Corps! Do you ladies understand?
[Full Metal Jacket]>>
<<I know why you're here, Neo. I know what you've been doing... why you hardly sleep, why you 
live alone, and why night after night, you sit by your computer. [...] I was looking for an 
answer. It's the question that drives us, Neo. It's the question that brought you here. You 
know the question, just as I did.
[The Matrix]>>
<<That system is our enemy. But when you're inside, you look around, what do you see? 
Businessmen, teachers, lawyers, carpenters. The very minds of the people we are trying to save. 
But until we do, these people are still a part of that system and that makes them our enemy. 
You have to understand, most of these people are not ready to be unplugged.
[The Matrix]>>
<<Fiery the angels fell. Deep thunder rolled around their shoulders... 
burning with the fires of Orc.
[Blade Runner]>>
<<There's no gene for fate.
[Gattaca]>>
<<Oh man, this isn't happening, it only thinks it's happening
[Tron]>>
<<Bazinga!
[Big Bang Theory]>>
<<If a tree falls in the forest and no one is around to hear it, what color is the tree?
[Monkey Island]>>
<<All science is either physics or stamp collecting
[Ernest Rutherford]>>
<<Stat rosa pristina nomine, nomina nuda tenemus
[Umberto Eco]>>
<<Theres a lady whos sure. All that glitters is gold. And shes buying a stairway to heaven.
[Led Zeppelin]>>
<<Django: This is the way things are. You can't change nature.
Remy: Change is nature, Dad. The part that we can influence. And it starts when we decide.
Django: [Remy turns to leave] Where are you going?
Remy: With luck, forward. 
[Ratatouille]>>
<<An idea is like a virus, resilient, highly contagious. The smallest seed of an idea can grow. 
It can grow to define or destroy you. 
[Inception]>>
<<How soft your fields so green, can whisper tales of gore,
Of how we calmed the tides of war. We are your overlords.
[Led Zeppelin]>>
<<I have to believe in a world outside my own mind. I have to believe that my actions still 
have meaning, even if I can't remember them. I have to believe that when my eyes are closed, 
the world's still there. Do I believe the world's still there? Is it still out there?... Yeah. 
We all need mirrors to remind ourselves who we are. I'm no different
[Memento]>>
<<The ice age is coming, the sun's zooming in
Meltdown expected, the wheat is growing thin
Engines stop running, but I have no fear
'Cause London is drowning, and I live by the river
[The Clash]>>
<<So, so you think you can tell Heaven from Hell, blue skies from pain.
Can you tell a green field from a cold steel rail? A smile from a veil?
[Pink Floyd]>>
<<Il corpo faccia quello che vuole. Io non sono il corpo: io sono la mente.
[Rita Levi Montalcini]>>
<<Considerate la vostra semenza
fatti non foste a viver come bruti
ma per seguir virtute e conoscenza
[Dante]>>
<<Fifteen men on the dead man's chest
Yo-ho-ho, and a bottle of rum!
[Robert Louis Stevenson]>>
<<Yoh ho, non c'e tregua nella gloria vivrà 
nel volto, vivo o morto, lei ti seguirà...
[hoist the colours - IT version]>>
<<Better sleep with a sober cannibal than a drunken Christian
[Hermann Melville]>>
<<When the man with a 45 meets the man with a rifle,
you said the man with a pistol is a dead man
[A fistful of dollars]>>
<<It lies unknown the land of mine
a hidden gate to save us from the shadow fall
the lord of water spoke
in the silence words of wisdom
I've seen the end of all
Be aware the storm gets closer
[Blind Guardian]>>
<<God of Rock, thank you for this chance to kick ass. We are your humble servants.
Please give us the power to blow people's minds with our high voltage rock. 
[School of rock]>>
<<Jake: First you traded the Cadillac in for a microphone. Then you lied to me 
about the band. And now you're gonna put me right back in the joint!
Elwood: They're not gonna catch us. We're on a mission from God.
[The Blues Brothers]>>
<<We are such stuff as dreams are made on;
and our little life is rounded with a sleep
[William Shakespeare]>>
<<Humpty Dumpty: When I use a word, it means just what I choose it to mean -
niether more or less
Alice: The questio is, whether you can make words mean so many different
things.
Humpty Dumpty: The question is, which is to be master - that's all
[Through the Looking-Glass]>>
#>