   M S B A M L   `   `   `  ����  SendBroadCastMessage"    StartConversation.MainWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$6  0Lync/Skype for Business - Broadcast Message Tool��5   	   $��514.147��6B   $	��710��6S   .��5      � $��	9,0,9,-73q�6    � ColumnDefinitions 5   
   �� 5      $
��262* �6   �� 5	      $
��298* �6   �� 5
      $
��115* �6   �� 5      $��1* �6   �� 5      $��1* �6   .+�5   
   �� $��Participants��6   $��0,0,-0.612,390q�6)   $
��Left=�6A   $	��675��6\   $��5��6h   .��5      �� -   5        $��rtbParticipants��5      $��1,1,1,1q�61    �� SelectionOpacity$ 0Q�6B   $��10��6W   $��1��6e    � 	Resources 5       �� 
TargetType �� Property �� Value��%X   'J�       �� 5      " J�6   .�5      ��  �@��Margin$ 0q�6/   5      5      5      5      �� 5   
   $��Status��6   $
��Left=�6#   $��0,415,-0.612,73q�6>   $	��675��6W   $��5��6c   .��5      �� -   5        $��	txtErrors��5       �� 
IsReadOnly . $��1,1,1,1q�69    J� RenderTransformOrigin$ 0.5,0.5&�6J   .��6k   5      5      �� 5   
   -   5        $��btnStartConv��5   T   $'��!1-Find Online Lync/Skype Contacts��6   $��2,380,240,141q�6=   $��2��6�   �� 5    
   -   5        $��btnSendMessage��5    _   $,��&2-Create Broadcast Conversation Window��6   $
��Left=�6u   $��
66,380,0,0q�6�   $	��Top=�6�   $	��341��6�   $��30��6�   $��2��6�   $��1��6�   �� 5"   
   $��Broadcast Message��6   $
��Left=�6.   $��0,171,-0.612,0q�6I   $	��Top=�6a   $��5��6y   $	��675��6�   $	��207��6�   .��5#      �� -   5        $��
rtbMessage��5#      $��1,1,1,1q�6.    5$      ��%X   'J�       �� 5%      " J�6   .�5&      ��  �@��Margin$ 0q�63   5'      5(      5)      5+      5,      5-      