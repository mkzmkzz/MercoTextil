"""
Test suite for MercoTêxtil system - NEW FEATURES:
1. Reset DB also clears banco_dados (artigos)
2. User permissions now include 'banco_dados' field
3. Export reports include engrenagem, enchimento, ciclos from artigos database
"""
import pytest
import requests
import os

BASE_URL = os.environ.get('REACT_APP_BACKEND_URL', '').rstrip('/')


@pytest.fixture(scope="module")
def auth_token():
    """Get auth token for tests"""
    response = requests.post(f"{BASE_URL}/api/auth/login", json={
        "username": "admin",
        "password": "admin123"
    })
    if response.status_code != 200:
        pytest.skip("Authentication failed - skipping tests")
    return response.json()["token"]


@pytest.fixture
def auth_headers(auth_token):
    """Auth headers fixture"""
    return {"Authorization": f"Bearer {auth_token}"}


class TestUserPermissionsBancoDados:
    """Tests for user permissions - banco_dados field"""
    
    def test_login_returns_banco_dados_permission(self, auth_headers):
        """GET /api/auth/me - Check user has banco_dados permission"""
        response = requests.get(f"{BASE_URL}/api/auth/me", headers=auth_headers)
        assert response.status_code == 200, f"Get me failed: {response.text}"
        
        user = response.json()
        assert "permissions" in user, "User should have permissions field"
        permissions = user["permissions"]
        
        # Verify banco_dados permission exists
        assert "banco_dados" in permissions, "banco_dados permission should exist in user permissions"
        # Admin should have banco_dados = True
        assert permissions["banco_dados"] == True, "Admin should have banco_dados permission"
    
    def test_create_user_with_banco_dados_permission(self, auth_headers):
        """POST /api/users - Create user with banco_dados permission"""
        import uuid
        unique_username = f"TEST_USER_{uuid.uuid4().hex[:8]}"
        
        user_data = {
            "username": unique_username,
            "email": f"{unique_username}@test.com",
            "password": "test123",
            "role": "operador_interno",
            "permissions": {
                "dashboard": True,
                "producao": True,
                "ordem_producao": True,
                "relatorios": True,
                "espulagem": True,
                "manutencao": True,
                "banco_dados": True,  # NEW: banco_dados permission
                "administracao": False
            }
        }
        
        response = requests.post(f"{BASE_URL}/api/users", json=user_data, headers=auth_headers)
        assert response.status_code == 200, f"Create user failed: {response.text}"
        
        created_user = response.json()
        assert created_user["permissions"]["banco_dados"] == True, "Created user should have banco_dados permission"
        
        # Cleanup - delete the test user
        user_id = created_user["id"]
        requests.delete(f"{BASE_URL}/api/users/{user_id}", headers=auth_headers)
    
    def test_update_user_banco_dados_permission(self, auth_headers):
        """PUT /api/users/{id} - Update user banco_dados permission"""
        import uuid
        unique_username = f"TEST_UPDATE_{uuid.uuid4().hex[:8]}"
        
        # Create user first
        user_data = {
            "username": unique_username,
            "email": f"{unique_username}@test.com",
            "password": "test123",
            "role": "operador_interno",
            "permissions": {
                "dashboard": True,
                "producao": True,
                "ordem_producao": True,
                "relatorios": True,
                "espulagem": True,
                "manutencao": True,
                "banco_dados": True,
                "administracao": False
            }
        }
        
        create_response = requests.post(f"{BASE_URL}/api/users", json=user_data, headers=auth_headers)
        assert create_response.status_code == 200
        user_id = create_response.json()["id"]
        
        # Update to disable banco_dados permission
        update_data = {
            "permissions": {
                "dashboard": True,
                "producao": True,
                "ordem_producao": True,
                "relatorios": True,
                "espulagem": True,
                "manutencao": True,
                "banco_dados": False,  # Change to False
                "administracao": False
            }
        }
        
        update_response = requests.put(f"{BASE_URL}/api/users/{user_id}", 
                                      json=update_data, headers=auth_headers)
        assert update_response.status_code == 200, f"Update user failed: {update_response.text}"
        
        updated_user = update_response.json()
        assert updated_user["permissions"]["banco_dados"] == False, "banco_dados should be updated to False"
        
        # Cleanup
        requests.delete(f"{BASE_URL}/api/users/{user_id}", headers=auth_headers)


class TestBancoDadosArtigos:
    """Tests for Banco de Dados - verifying artigos have all required fields"""
    
    def test_create_artigo_with_all_fields(self, auth_headers):
        """POST /api/banco-dados - Create artigo with engrenagem, ciclos, carga"""
        artigo_data = {
            "artigo": "TEST_ARTIGO_FULL",
            "engrenagem": "ENG-TEST-100",
            "fios": "24",
            "maquinas": "CD1, CD2, CD3",
            "ciclos": "5",
            "carga": "100kg"
        }
        
        response = requests.post(f"{BASE_URL}/api/banco-dados", json=artigo_data, headers=auth_headers)
        assert response.status_code == 200, f"Create artigo failed: {response.text}"
        
        artigo = response.json()
        assert artigo["artigo"] == "TEST_ARTIGO_FULL"
        assert artigo["engrenagem"] == "ENG-TEST-100"
        assert artigo["ciclos"] == "5"
        assert artigo["carga"] == "100kg"
        
        # Cleanup
        requests.delete(f"{BASE_URL}/api/banco-dados/{artigo['id']}", headers=auth_headers)
    
    def test_search_artigo_returns_all_fields(self, auth_headers):
        """GET /api/banco-dados/search - Verify search returns engrenagem, ciclos, carga"""
        # Create artigo first
        artigo_data = {
            "artigo": "TEST_SEARCH_ARTIGO",
            "engrenagem": "ENG-SEARCH",
            "fios": "32",
            "maquinas": "F1, F2",
            "ciclos": "3",
            "carga": "50kg"
        }
        
        create_response = requests.post(f"{BASE_URL}/api/banco-dados", json=artigo_data, headers=auth_headers)
        assert create_response.status_code == 200
        artigo_id = create_response.json()["id"]
        
        # Search for it
        search_response = requests.get(f"{BASE_URL}/api/banco-dados/search?q=TEST_SEARCH", headers=auth_headers)
        assert search_response.status_code == 200
        
        results = search_response.json()
        assert len(results) >= 1
        
        # Find our artigo
        found_artigo = None
        for a in results:
            if a["artigo"] == "TEST_SEARCH_ARTIGO":
                found_artigo = a
                break
        
        assert found_artigo is not None, "Search should return our test artigo"
        assert found_artigo["engrenagem"] == "ENG-SEARCH"
        assert found_artigo["ciclos"] == "3"
        assert found_artigo["carga"] == "50kg"
        
        # Cleanup
        requests.delete(f"{BASE_URL}/api/banco-dados/{artigo_id}", headers=auth_headers)


class TestExportReportEndpoints:
    """Tests for export endpoints - verify they return data for espulagem reports"""
    
    def test_espulas_endpoint_returns_data(self, auth_headers):
        """GET /api/espulas - Verify espulas endpoint works"""
        response = requests.get(f"{BASE_URL}/api/espulas", headers=auth_headers)
        assert response.status_code == 200, f"Get espulas failed: {response.text}"
        
        data = response.json()
        assert isinstance(data, list)
    
    def test_banco_dados_endpoint_for_export(self, auth_headers):
        """GET /api/banco-dados - Verify banco-dados returns all fields needed for export"""
        response = requests.get(f"{BASE_URL}/api/banco-dados", headers=auth_headers)
        assert response.status_code == 200, f"Get banco-dados failed: {response.text}"
        
        data = response.json()
        assert isinstance(data, list)
        
        # If there are artigos, verify they have the fields needed for export
        if len(data) > 0:
            artigo = data[0]
            # These fields should exist (may be empty strings but should exist)
            assert "artigo" in artigo
            assert "engrenagem" in artigo
            assert "fios" in artigo
            assert "maquinas" in artigo


class TestInitDataFunction:
    """Tests for init_data - verify banco_dados is included"""
    
    def test_admin_user_has_banco_dados_permission(self, auth_headers):
        """Verify admin user created by init_data has banco_dados permission"""
        response = requests.get(f"{BASE_URL}/api/auth/me", headers=auth_headers)
        assert response.status_code == 200
        
        user = response.json()
        # Admin should have banco_dados permission set to True
        assert user.get("permissions", {}).get("banco_dados") == True


class TestDefaultUsersPermissions:
    """Test default user permissions include banco_dados"""
    
    def test_get_users_default_users_have_banco_dados_field(self, auth_headers):
        """GET /api/users - Verify default users (admin, interno, externo) have banco_dados in permissions"""
        response = requests.get(f"{BASE_URL}/api/users", headers=auth_headers)
        assert response.status_code == 200, f"Get users failed: {response.text}"
        
        users = response.json()
        assert len(users) > 0, "Should have at least one user"
        
        # Check only default users created by init_data
        default_usernames = ["admin", "interno", "externo"]
        
        for user in users:
            if user.get("username") in default_usernames and "permissions" in user:
                # banco_dados should be in permissions for default users
                assert "banco_dados" in user["permissions"], f"Default user {user['username']} should have banco_dados permission field"
                print(f"User {user['username']} has banco_dados permission: {user['permissions'].get('banco_dados')}")



class TestMapaTrancadeirasEndpoint:
    """Tests for GET /api/reports/mapa-trancadeiras"""

    def test_get_mapa_trancadeiras_32_fusos(self, auth_headers):
        """Returns machines list with order data for 32 fusos layout"""
        response = requests.get(
            f"{BASE_URL}/api/reports/mapa-trancadeiras?layout_type=32_fusos",
            headers=auth_headers
        )
        assert response.status_code == 200, f"Request failed: {response.text}"
        data = response.json()
        assert "machines" in data
        assert "generated_at" in data
        assert data["layout_type"] == "32_fusos"
        assert len(data["machines"]) > 0
        for machine in data["machines"]:
            assert "code" in machine
            assert "status" in machine
            assert "order" in machine

    def test_get_mapa_trancadeiras_16_fusos(self, auth_headers):
        """Returns machines list with order data for 16 fusos layout"""
        response = requests.get(
            f"{BASE_URL}/api/reports/mapa-trancadeiras?layout_type=16_fusos",
            headers=auth_headers
        )
        assert response.status_code == 200, f"Request failed: {response.text}"
        data = response.json()
        assert "machines" in data
        assert data["layout_type"] == "16_fusos"
        assert len(data["machines"]) > 0

    def test_get_mapa_trancadeiras_no_auth(self):
        """Unauthenticated request should be rejected"""
        response = requests.get(
            f"{BASE_URL}/api/reports/mapa-trancadeiras?layout_type=32_fusos"
        )
        assert response.status_code in [401, 403]

    def test_get_mapa_trancadeiras_externo_forbidden(self):
        """operador_externo should NOT access this endpoint"""
        login_response = requests.post(f"{BASE_URL}/api/auth/login", json={
            "username": "externo",
            "password": "externo123"
        })
        if login_response.status_code != 200:
            pytest.skip("External user login failed")
        externo_token = login_response.json()["token"]
        response = requests.get(
            f"{BASE_URL}/api/reports/mapa-trancadeiras?layout_type=32_fusos",
            headers={"Authorization": f"Bearer {externo_token}"}
        )
        assert response.status_code == 403

    def test_get_mapa_trancadeiras_interno_allowed(self):
        """operador_interno should be able to access this endpoint"""
        login_response = requests.post(f"{BASE_URL}/api/auth/login", json={
            "username": "interno",
            "password": "interno123"
        })
        if login_response.status_code != 200:
            pytest.skip("Internal user login failed")
        interno_token = login_response.json()["token"]
        response = requests.get(
            f"{BASE_URL}/api/reports/mapa-trancadeiras?layout_type=32_fusos",
            headers={"Authorization": f"Bearer {interno_token}"}
        )
        assert response.status_code == 200


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
