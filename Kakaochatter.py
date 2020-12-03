import time, ctypes
import win32con as wc
import win32api as wa
import win32gui as wg
from pywinauto import clipboard
from datetime import datetime
import re


class Kakaotime():
    def now(self):
        tm = datetime.now().time()
        if tm.hour > 12:
            tm = '오후 {0}:{1:02d}'.format(tm.hour-12, tm.minute)
        else:
            tm = '오전 {0}:{1:02d}'.format(tm.hour, tm.minute)
        return tm
    
    def convert(self, hour, min):
        if isinstance(hour, int): raise TypeError('인자 hour은 str 객체여야 합니다.')
        if isinstance(min, int): raise TypeError('인자 min은 str 객체여야 합니다.')

        tm = ''
        if hour > 12:
            tm = '오후 {0}:{1:02d}'.format(hour-12, minute)
        else:
            tm = '오전 {0}:{1:02d}'.format(hour, minute)
        return tm


class Kakaochat():
    def __init__(self, name, tm, text):
        self.set(name, tm, text)

    def __str__(self):
        return f'[{self.name}] [{self.tm}] {self.text}'

    def set(self, name, tm, text):
        if not isinstance(name, str): raise TypeError('인자 name은 str 객체여야 합니다.')
        if not isinstance(tm, str): raise TypeError('인자 tm은 str 객체여야 합니다.')
        if not isinstance(text, str): raise TypeError('인자 text은 str 객체여야 합니다.')

        self.name = name
        self.tm = tm
        self.text = text

    def __eq__(self, other):
        return (self.name==other.name and self.tm==other.tm and self.text==other.text)


class KakaochatList():
    def __init__(self, kcl=None, s=0, e=-1):
        if kcl and not isinstance(kcl, KakaochatList): raise TypeError('인자 kcl은 Kakaochat 객체여야 합니다.')
        if not isinstance(s, int): raise TypeError('인자 s는 정수여야 합니다.')
        if not isinstance(s, int): raise TypeError('인자 e는 정수여야 합니다.')

        if kcl:
            if e == -1:
                self.chatlist = list(kcl.chatlist[s:])
            else:
                self.chatlist = list(kcl.chatlist[s:e])
        else:
            self.chatlist = list()
    
    def add(self, kakaochat):
        if isinstance(kakaochat, Kakaochat):
            self.chatlist.append(kakaochat)
        else:
            raise TypeError('인자 kakaochat은 Kakaochat 객체여야 합니다.')
    
    def fromlast(self, index):
        if isinstance(index, int):
            return self.chatlist[-index-1]
        else:
            raise TypeError('인자 index는 정수여야 합니다.')

    
    def __str__(self):
        ret = str()
        for c in self.chatlist:
            ret += c.getasText() + '\n'
        return ret

    def index(self, fkc):
        if isinstance(fkc, Kakaochat):
            return self.chatlist.index(fkc)
        else:
            raise TypeError('인자 kakaochat은 Kakaochat 객체여야 합니다.')

    def indexfromlast(self, fkc):
        if isinstance(fkc, Kakaochat):
            return len(self.chatlist)-1 - self.chatlist[::-1].index(fkc)
        else:
            raise TypeError('인자 kakaochat은 Kakaochat 객체여야 합니다.')


class KakaoChatter():
    def __init__(self, myname, chatroom_name):
        if isinstance(myname, str): raise TypeError('인자 myname은 str 객체여야 합니다.')
        if isinstance(chatroom_name, str): raise TypeError('인자 chatroom_name은 str 객체여야 합니다.')

        self.myname = myname
        self.setNewChatroom(chatroom_name)
        self._user32 = ctypes.WinDLL("user32")
        self.lastchat = None
        
    def setNewChatroom(self, chatroom_name):
        if isinstance(chatroom_name, str): raise TypeError('인자 chatroom_name은 str 객체여야 합니다.')
        
        self.chatroom_name = chatroom_name
        self.hwndMain = wg.FindWindow(None, self.chatroom_name)
        if not wg.isWindow(self.hwndMain):
            open_chatroom()
        self.hwndEdit = wg.FindWindowEx(self.hwndMain, None, "RICHEDIT50W", None)
        self.hwndList = wg.FindWindowEx(self.hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    def sendReturn(self):
        wa.PostMessage(self.hwndEdit, wc.WM_KEYDOWN, wc.VK_RETURN, 0)
        time.sleep(0.01)
        wa.PostMessage(self.hwndEdit, wc.WM_KEYUP, wc.VK_RETURN, 0)
        
    def sendText(self, text):
        if isinstance(text, str): raise TypeError('인자 text은 str 객체여야 합니다.')
        wa.SendMessage(self.hwndEdit, wc.WM_SETTEXT, 0, text)
        self.sendReturn()
        self.setLastchat(self.readLastchat())

    def setChatroom_open(self, chatroom_name):
        if isinstance(chatroom_name, str): raise TypeError('인자 chatroom_name은 str 객체여야 합니다.')

        # set chatroom open
        pass

    def open_chatroom(self):
        hwndkakao = wg.FindWindow(None, "카카오톡")
        hwndkakao_edit1 = wg.FindWindowEx(hwndkakao, None, "EVA_ChildWindow", None)
        hwndkakao_edit2_1 = wg.FindWindowEx(hwndkakao_edit1, None, "EVA_Window", None)
        hwndkakao_edit2_2 = wg.FindWindowEx(hwndkakao_edit1, hwndkakao_edit2_1, "EVA_Window", None)
        hwndkakao_edit3 = wg.FindWindowEx(hwndkakao_edit2_2, None, "Edit", None)

        wa.SendMessage(hwndkakao_edit3, wc.WM_SETTEXT, 0, self.chatroom_name)
        time.sleep(1)   # 안정성 위해 필요
        SendReturn(hwndkakao_edit3)
        time.sleep(1)

        
    def readchat(self):
        self.PostVirtualKey(ord('A'), [wc.VK_CONTROL])
        time.sleep(0.1)
        self.PostVirtualKey(ord('C'), [wc.VK_CONTROL])
        lc = re.split('\r|\n', clipboard.GetData())
        
        ret = KakaochatList()
        for line in lc:
            try:
                name = line[1:line.index(']')]
                line = line[line.index(']')+3:]
                tm = line[:line.index(']')]
                line = line[line.index(']')+2:]
                ret.add(Kakaochat(name, tm, line))
            except:
                pass

        return ret

    def readNewchat(self):
        cl = self.readchat()
        newchat_list = KakaochatList(cl, cl.indexfromlast(self.lastchat))
        self.lastchat = newchat_list.fromlast(0)
        return newchat_list

    def readLastchat(self):
        cl = self.readchat()
        return cl.fromlast(0)
    
    def setLastchat(self, kc):
        if not isinstance(kc, Kakaochat):
            raise TypeError('인자 kc은 Kakaochat 객체여야 합니다.')
        self.lastchat = kc
        
    def PostVirtualKey(self, key, shift):
        # https://whynhow.info/44338/How-to-transfer-the-key-combination-(CTRL--A-etc)-to-an-inactive-window?
        PBYTE256 = ctypes.c_ubyte * 256
        GetKeyboardState = self._user32.GetKeyboardState
        SetKeyboardState = self._user32.SetKeyboardState
        GetWindowThreadProcessId = self._user32.GetWindowThreadProcessId
        AttachThreadInput = self._user32.AttachThreadInput

        MapVirtualKeyA = self._user32.MapVirtualKeyA

        if wg.IsWindow(self.hwndList):
            ThreadId = GetWindowThreadProcessId(self.hwndList, None)

            lparam = wa.MAKELONG(0, MapVirtualKeyA(key, 0))

            if len(shift) > 0:  # Если есть модификаторы - используем PostMessage и AttachThreadInput
                pKeyBuffers = PBYTE256()
                pKeyBuffers_old = PBYTE256()

                wa.SendMessage(self.hwndList, wc.WM_ACTIVATE, wc.WA_ACTIVE, 0)
                AttachThreadInput(wa.GetCurrentThreadId(), ThreadId, True)
                GetKeyboardState(ctypes.byref(pKeyBuffers_old))

                for modkey in shift:
                    if modkey == wc.VK_MENU:
                        lparam = lparam | 0x20000000
                        msg_down = wc.WM_SYSKEYDOWN
                        msg_up = wc.WM_SYSKEYUP
                    pKeyBuffers[modkey] |= 128

                SetKeyboardState(ctypes.byref(pKeyBuffers))
                time.sleep(0.01)
                wa.PostMessage(self.hwndList, wc.WM_KEYDOWN, key, lparam)
                time.sleep(0.01)
                wa.PostMessage(self.hwndList, wc.WM_KEYUP, key, lparam | 0xC0000000)
                time.sleep(0.01)
                SetKeyboardState(ctypes.byref(pKeyBuffers_old))
                time.sleep(0.01)
                AttachThreadInput(wa.GetCurrentThreadId(), ThreadId, False)

            else:  # Если нету модификаторов - используем SendMessage
                SendMessage(self.hwndList, msg_down, key, lparam)
                SendMessage(self.hwndList, msg_up, key, lparam | 0xC0000000)

            return True
        
        else:
            return False


def main2():
    kc = KakaoChatter('김영후', r'김영후')
    #kc = KakaoChatter('뉴비', r'카카오톡 봇 만들기 방&카카오톡 봇 질문방')
    #kc = KakaoChatter('강인공지능', r'Abstract')
    cog = 'ㄸ'
    startmsg = f'앵무새봇 시작. {cog}뒤에 오는 문장을 따라합니다. {cog}종료를 입력하면 종료합니다.'
    endmsg = f'{cog}종료'
    readcycle = 0.5
    kc.sendText(startmsg)
    goodend = False
    while wg.IsWindow(kc.hwndMain) and (not goodend):
        newchat = kc.readNewchat()
        for nc in newchat.chatlist:
            if nc.text == endmsg:
                goodend = True
                break

            if len(nc.text) > len(cog) and nc.text[:len(cog)] == 'ㄸ':
                kc.sendText(nc.text[len(cog):])

            time.sleep(0.2)

        time.sleep(readcycle)

    if goodend:
        kc.sendText('앵무새봇이 정상적으로 종료되었습니다.')
    else:
        print('앵무새봇이 강제 종료되었습니다.')



def main():
    kc = KakaoChatter('뉴비', '연두님이 고물취급하는 봇')
    runtime = 10
    readcycle = 0.5
    startmessage = f'앵무새봇 시작 약 {runtime}초 후에 종료'
    kc.sendText(startmessage)
    prev = startmessage
    print(kc.readLastchat())
    for i in range(int(runtime/readcycle)):
        lastchat = kc.readLastchat()[2]
        if prev != lastchat:
            kc.sendText(lastchat)
        prev = lastchat
        time.sleep(readcycle)

    kc.sendText('앵무새봇 종료')


if __name__ == '__main__':
    main2()
