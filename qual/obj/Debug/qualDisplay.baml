   M S B A M L   `   `   `  ����?  ;qual, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:qualqual  5          qual.qualDisplay   5      R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      $localclr-namespace:qual     5        N� Title$  Qual Report - View Report��5	      .��5      � $��
10,0,10,10q�6    � ColumnDefinitions 5   
   �� 5       � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      .+�5   
   � $��0��6    5      �� 5      $
��Auto �6     5      �� 5      $
��Auto �6#   �� 5      $
��Auto �6#   �� 5      $
��Auto �6#   �� 5      $
��Auto �6#   5!      &   System.Windows.Controls.DataGrid  5"   
   -   5        $��displayGrid��5"       $��3��6      AutoGenerateColumns .     Columns 5#      0  *System.Windows.Controls.DataGridTextColumn  5$      ,  &System.Windows.Controls.DataGridColumn   Header$ ID��6%   1  +System.Windows.Controls.DataGridBoundColumn   Binding 61   �� *ID+  5%      $ 	Last Name��6%    68   �� *
LastName+  5&      $ 
First Name��6%    69   �� *	GivenName+  5'      $ Certification��6%    6<   �� *Cert+  5(      $ Expiry Date��6%    6:   �� *Expiry+  5)      $ 	Days Left��6%    68   �� *
DaysLeft+4  .System.Windows.Controls.DataGridCheckBoxColumn  5*      $ Email Sent?��6)    6>   �� *	Emailed+  5+      $ Sent on��6%    66   �� *EmailedDate+4  .System.Windows.Controls.DataGridTemplateColumn  5,      $ 
Send Email��6)      CellTemplate 5-      �� 5.      .R�5/      �� $��btnSend��6%   -   5        $
��Send��5/   4   6C   50      52      54      55      56      