   M S B A M L   `   `   `  ����  QuickTimeEntry&     QuickTimeEntry.SelectMSCRMEntity   5   	   f bDevExpress.Xpf.Grid.v16.1.Core, Version=16.1.4.0, Culture=neutral, PublicKeyToken=b88d1754d700e49aa ]DevExpress.Xpf.Grid.v16.1, Version=16.1.4.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a>dxg2http://schemas.devexpress.com/winfx/2008/xaml/grid   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35c _Infralution.Localization.Wpf, Version=3.3.0.0, Culture=neutral, PublicKeyToken=547ccae517a004b5a ]DevExpress.Xpf.Core.v16.1, Version=16.1.4.0, Culture=neutral, PublicKeyToken=b88d1754d700e49aW SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   ;dx2http://schemas.devexpress.com/winfx/2008/xaml/core  5   	   ?dxe5http://schemas.devexpress.com/winfx/2008/xaml/editors  5   7   0  *Infralution.Localization.Wpf.ResxExtension    DefaultResxName$&   QuickTimeEntry.SelectMSCRMEntity��5      $	��500��5      $	��550��6+    � 
ResizeMode$ NoResize=�6y    � ShowInTaskbar .   � Topmost .  � WindowStartupLocation$ CenterOwner=�6�    � WindowStyle$ 
ToolWindow=�6�    � Language 5   D   5  /Infralution.Localization.Wpf.UICultureExtension   � Title 5   	     *Title+.��5      � &   DevExpress.Xpf.Core.ThemeManager   	ThemeName0  *DevExpress.Xpf.Core.ThemeNameTypeConverter$ Office2010Silver 6   $
��Auto��68   $
��Auto��6E   $��5,5,5,5q�6S   	 � RowDefinitions	 5	   
   �� 5
      $��30 �6   �� 5      $��* �6   .+�5   
   � $��0��6   
 � ColumnDefinitions
 5      �� 5      $��* �6#   �� 5      $	��100 �6#   	 5      �� 5      $��35 �6    .+�5      )  #DevExpress.Xpf.Editors.ComboBoxEdit  -   5        $��tbEntitySearch��5   m   $��5,5,5,5q�6   
���   �$��0��6D   $��0��6T   $��23��6a   $
��Auto��6�   �� 5      -   5        $��	btnSearch��5   T    �� 	IsDefault . $��5,5,5,5q�6&   $��1��67   $��0��6G   $��90��6e   $��23��6p   .��5      ��  �� Orientation$ 
Horizontal=�6   .+�5      �� -   5        $��btnSearchImage��5   '   $��16��6   $��0,0,5,0q�6=   $:��4/QuickTimeEntry;component/Images/16x16neu/Suchen.png��6N   �� 5      ��6      *	btnSearch+6:   5      5      5      � 5   
   $��1��6   
 5      �� 5      $
��357* �6#   �� 5       $
��133* �6#   �� 5!      $	��44* �6#   	 5#      �� 5$      $��* �6    �� 5%      $��30 �6    .+�5'      %  DevExpress.Xpf.Grid.GridControl  -   5        $��
SearchGrid��5'   D   $��5q�6   $
��Auto��6)   $
��Auto��67   $��3��6�      View 5(      #  DevExpress.Xpf.Grid.TableView  5)      -   5          Name$ SearchGridview��5)   $      	AutoWidth . &   DevExpress.Xpf.Grid.DataViewBase   NavigationStyle$
 Cell=�6M      ShowTotalSummary . 6d      Columns 5+      $	  DevExpress.Xpf.Grid.GridColumn	  5,      -   5        $
  DevExpress.Xpf.Grid.BaseColumn 
 Name$ gc_id��5,   %   $  DevExpress.Xpf.Grid.ColumnBase   	FieldName$ id��6H    
  Visible .   
  Header 62     *gc_id+6W   	  5-      -   5        $ gc_name��5-   %   $
 name��6L    64     *	gc_name+6L   5/      � 50      $��1��6   $��3��6    
 51      �� 52      $��* �6'   �� 53      $	��160 �6'   �� 54      $	��115 �6'   .+�56      �� -   5        $��
searchInfo��56      $��0��6*   $��5,5,0,0q�6:   $�� ��6L   �� 57      -	   5        $��	btnAccept��57   F   $��5,0,5,5q�6   $��1��6*   $��23��6:   $	��150��6W   .��58      �� $ 
Horizontal=�6!   .+�59      �� -
   5        $��btnAcceptImage��59   +   $��16��6    $��0,0,5,0q�6A   $=��7/QuickTimeEntry;component/Images/16x16neu/Speichern.png��6R   �� 5:      ��6$     *	btnAccept+6>   5;      5<      �� 5=      -   5        $��btnClose��5=   F   $��5,0,5,5q�6   $��2��6*   $��23��6:   $	��105��6V   .��5>      �� $ 
Horizontal=�6!   .+�5?      �� -   5        $��btnCancelTimeentryImage_1��5?   +   $��16��6    $��0,0,5,0q�6L   $=��7/QuickTimeEntry;component/Images/16x16neu/Abbrechen.png��6]   �� 5@      ��6$     *
btnClose+6=   5A      5B      5C      5D      5E      5F      