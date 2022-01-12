# visualbasic
Sourcecode for Visual Basic

RC4.bas is a Implementation of the RC4 and the much stronger Spritz algorithm, see: https://www.schneier.com/blog/archives/2014/10/spritz_a_new_rc.html based on this RC4 implementation: https://github.com/g2jun/RC4-VB

=> It is possible to switch between RC4 (False) and Spritz (True) with the blSpritz switch.

=> With the RC4_Crypt_test() it is shown that with the same settings a plaintext can be encrypted and decrypted again. For this the Spritz algorithm was used and  bytW / w = 99 was set (according to Bruce Schneier this can be any odd number between 1 and 255).
