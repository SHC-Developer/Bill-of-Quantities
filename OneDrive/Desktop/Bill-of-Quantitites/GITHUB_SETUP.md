# GitHub 푸시 가이드

이 프로젝트를 GitHub에 푸시하기 위한 단계별 가이드입니다.

## 1. Git 저장소 초기화 (아직 안 했다면)

```bash
git init
```

## 2. 모든 파일 추가

```bash
git add .
```

## 3. 첫 커밋 생성

```bash
git commit -m "Initial commit: 외관조사망도 도면별 물량표 추출 사이트"
```

## 4. 원격 저장소 추가

```bash
git remote add origin https://github.com/SHC-Developer/Bill-of-Quantities.git
```

## 5. 기본 브랜치를 main으로 설정 (필요한 경우)

```bash
git branch -M main
```

## 6. GitHub에 푸시

```bash
git push -u origin main
```

또는 master 브랜치를 사용하는 경우:

```bash
git push -u origin master
```

## 주의사항

- GitHub 저장소가 이미 생성되어 있어야 합니다
- GitHub 인증이 필요할 수 있습니다 (Personal Access Token 또는 SSH 키)
- 처음 푸시하는 경우 `-u` 옵션을 사용하여 upstream을 설정합니다

## 문제 해결

### 인증 오류가 발생하는 경우

1. Personal Access Token 사용:
   - GitHub Settings > Developer settings > Personal access tokens
   - 토큰 생성 후 비밀번호 대신 사용

2. 또는 SSH 키 사용:
   ```bash
   git remote set-url origin git@github.com:SHC-Developer/Bill-of-Quantities.git
   ```

